# excel_processor/processor.py
import time
import pandas as pd
import xlwings as xw
from typing import List, Tuple, Optional, Dict

from .models import ProcessingConfig, ProcessingResult
from .com_management import COMManager, EnhancedExcelOptimizer
from .subsidiary import SubsidiaryExtractor
from .memory_optimizer import MemoryOptimizer

# Import the new modules
from .exceptions import *
from .retry_utils import safe_excel_operation_with_retry, retry_on_failure, RetryConfig
from .performance_logger import performance_logger, log_performance

class EnhancedExcelProcessor:
    def __init__(self, config: ProcessingConfig):
        self.config = config
        self.summary_data: Optional[pd.DataFrame] = None
        self.summary_lookup: Dict[str, tuple] = {}
        self.subsidiary_variations: Dict[str, str] = {}

    # ---------- SUMMARY ----------
    def load_summary_data_enhanced(self, summary_path: str):
        print("📊 Loading and analyzing summary data...")
        self.summary_data = pd.read_excel(summary_path, dtype=str).fillna('')

        subsidiaries = self.summary_data['Subsidiary'].unique()
        for sub in subsidiaries:
            if pd.notna(sub) and sub.strip():
                clean = sub.strip().upper()
                self.subsidiary_variations[clean] = sub
                if '-' in clean:
                    self.subsidiary_variations[clean.split('-')[0].strip()] = sub

        self.summary_lookup = {}
        for idx, row in self.summary_data.iterrows():
            k1 = f"{row['Unit name'].strip()}|{row['Tenant ID'].strip()}"
            k2 = f"{row['Unit name'].strip()}|{row['Tenant'].strip()}"
            self.summary_lookup[k1] = (idx, row.to_dict())
            self.summary_lookup[k2] = (idx, row.to_dict())

        print(f"   ✅ Loaded {len(self.summary_data)} summary records")
        print(f"   ✅ Created {len(self.summary_lookup)} lookup keys")

    def get_subsidiary_subset(self, extracted_subsidiary: str) -> pd.DataFrame:
        if not extracted_subsidiary:
            return self.summary_data
        ss = self.summary_data
        exact = ss[ss['Subsidiary'].astype(str).str.strip().str.upper() == extracted_subsidiary.upper()]
        if not exact.empty:
            return exact
        for var, original in self.subsidiary_variations.items():
            if var == extracted_subsidiary.upper():
                m = ss[ss['Subsidiary'].astype(str).str.strip() == original]
                if not m.empty:
                    print(f"   🔄 Matched {extracted_subsidiary} -> {original}")
                    return m
        partial = ss[ss['Subsidiary'].astype(str).str.contains(extracted_subsidiary, case=False, na=False)]
        if not partial.empty:
            print(f"   🔍 Partial match for {extracted_subsidiary}")
            return partial
        print(f"   ⚠️ No subsidiary match for '{extracted_subsidiary}'")
        return pd.DataFrame()

    # ---------- CORE PER-FILE ----------
    @log_performance("process_single_file")
    def process_single_file_enhanced(self, filepath: str) -> ProcessingResult:
        start = time.time()
        result = ProcessingResult(filepath=filepath, status='error')

        app: Optional[xw.App] = None
        wb = None
        try:
            if not COMManager.initialize_com():
                result.error_message = "COM initialization failed"
                return result

            print(f"\n🔄 Processing: {filepath}")
            # Use connection pooling for better resource management
            app_id = f"worker_{hash(filepath) % 100}"  # Distribute files across app pool
            app = COMManager.get_or_create_excel_app(app_id)
            if not app:
                result.error_message = "Could not initialize Excel application"
                return result

            wb = safe_excel_operation_with_retry(
                lambda: app.books.open(filepath), 
                "Open workbook", 
                max_attempts=3
            )
            
            # Apply memory optimizations for large files
            MemoryOptimizer.optimize_workbook_for_large_files(wb)

            # chọn sheet
            try:
                sheet = wb.sheets['1.Leasing income']
            except Exception:
                names = [s.name for s in wb.sheets]
                candidates = [n for n in names if 'leasing' in n.lower() or 'income' in n.lower()]
                if candidates:
                    sheet = wb.sheets[candidates[0]]
                    print(f"   📋 Using sheet: {candidates[0]}")
                else:
                    raise Exception(f"Leasing income sheet not found. Available: {names}")

            header_row = EnhancedExcelOptimizer.find_header_row_enhanced(sheet)
            if not header_row:
                result.error_message = "Header row not found"
                wb.close()
                return result

            subsidiary = SubsidiaryExtractor.extract_subsidiary_enhanced(sheet, filepath, header_row)
            result.subsidiary_found = subsidiary

            summary_subset = self.get_subsidiary_subset(subsidiary)
            result.summary_matches = len(summary_subset)
            if summary_subset.empty:
                result.error_message = f"No summary data for subsidiary '{subsidiary}'"
                wb.close()
                return result

            print("   📊 Reading sheet data...")
            start_memory = MemoryOptimizer.get_memory_usage()
            headers, data = self._batch_read_enhanced(sheet, header_row)
            end_memory = MemoryOptimizer.get_memory_usage()
            print(f"   📈 Data read completed: {end_memory - start_memory:+.1f}MB memory change")
            if not data:
                result.error_message = "No data rows found"
                wb.close()
                return result

            df = pd.DataFrame(data, columns=headers).astype(object).fillna('')
            rows_updated, rows_added = self._process_dataframe_enhanced(df, sheet, header_row, headers, summary_subset)

            print("   💾 Saving workbook...")
            wb.save()
            wb.close()
            MemoryOptimizer.cleanup_memory()

            result.status = 'success'
            result.rows_updated = rows_updated
            result.rows_added = rows_added
            result.processing_time = time.time() - start
            print(f"   ✅ Success: {rows_updated} updated, {rows_added} added ({result.processing_time:.1f}s)")

        except Exception as e:
            result.error_message = str(e)
            print(f"   ❌ Error: {e}")
            try:
                if wb: wb.close()
            except: pass
        finally:
            try:
                # Release app back to pool instead of quitting
                if app and 'app_id' in locals():
                    COMManager.release_excel_app(app_id)
                else:
                    # Fallback cleanup for apps not from pool
                    if app: app.quit()
            except Exception as e:
                print(f"   ⚠️ Excel cleanup warning: {e}")
            time.sleep(0.1)  # Reduced cleanup delay
        return result

    # ---------- IO helpers ----------
    def _batch_read_enhanced(self, sheet: xw.Sheet, header_row: int) -> Tuple[List[str], List[List]]:
        """
        Enhanced batch reading for large files with dynamic size calculation
        Removed artificial 300-row, 40-column limits for 20MB file support
        """
        try:
            # Get actual used range without artificial limits
            used = EnhancedExcelOptimizer.safe_excel_operation(lambda: sheet.used_range)
            last_cell = EnhancedExcelOptimizer.safe_excel_operation(lambda: used.last_cell)
            actual_last_row = int(last_cell.row)
            actual_last_col = int(last_cell.column)
            
            print(f"   📊 Full used range detected: Row {header_row}..{actual_last_row}, Col 1..{actual_last_col}")
            
            # Calculate dynamic limits based on available memory
            available_memory_mb = MemoryOptimizer.get_available_memory()
            estimated_cell_size_bytes = 50  # Average bytes per cell (text/number)
            max_cells_for_memory = int((available_memory_mb * 0.3 * 1024 * 1024) / estimated_cell_size_bytes)
            
            # Dynamic row/column calculation
            if actual_last_col <= 50:  # Small width files
                max_rows_for_memory = min(max_cells_for_memory // actual_last_col, 100000)
            else:  # Wide files
                max_rows_for_memory = min(max_cells_for_memory // actual_last_col, 50000)
                
            # Use actual size but respect memory limits
            last_row = min(actual_last_row, max_rows_for_memory + header_row)
            last_col = actual_last_col
            
            if last_row < actual_last_row:
                print(f"   ⚠️ Memory constraint: Processing {last_row - header_row} of {actual_last_row - header_row} data rows")
            else:
                print(f"   ✅ Processing all {actual_last_row - header_row} data rows")
                
        except Exception as e:
            # Fallback to reasonable defaults if range detection fails
            print(f"   ⚠️ Could not determine used range: {e}")
            last_row, last_col = header_row + 10000, 100  # Much higher fallback limits
            print(f"   🔄 Using fallback limits: Row {header_row}..{last_row}, Col 1..{last_col}")

        # Read headers with dynamic column range
        headers_raw = EnhancedExcelOptimizer.safe_excel_operation(
            lambda: sheet.range((header_row, 1), (header_row, last_col)).value
        )
        headers = [str(h).strip() if h else f'Col_{i}' for i, h in enumerate(headers_raw)]

        # Handle duplicate Rent columns
        rent_idx = [i for i, h in enumerate(headers) if h == 'Rent']
        if len(rent_idx) >= 2:
            headers[rent_idx[0]] = 'Rent (USD)'
            headers[rent_idx[1]] = 'Rent (VND)'
            print("   🔄 Renamed duplicate Rent columns")

        data = []
        if last_row > header_row:
            total_rows = last_row - header_row
            
            # Use chunked reading for large datasets
            if total_rows > self.config.chunk_size:
                print(f"   📦 Using chunked reading: {self.config.chunk_size} rows per chunk")
                data = self._read_data_chunked(sheet, header_row + 1, last_row, last_col)
            else:
                # Single read for smaller datasets
                try:
                    print(f"   ⚡ Reading {total_rows} rows at once...")
                    all_data = EnhancedExcelOptimizer.safe_excel_operation(
                        lambda: sheet.range((header_row + 1, 1), (last_row, last_col)).value
                    )
                    if all_data:
                        if not isinstance(all_data, list):
                            data = [all_data]
                        elif len(all_data) > 0 and not isinstance(all_data[0], list):
                            data = [all_data]
                        else:
                            data = all_data
                except Exception as e:
                    print(f"   ⚠️ Bulk read failed, falling back to chunked read: {e}")
                    data = self._read_data_chunked(sheet, header_row + 1, last_row, last_col)
        
        print(f"   📚 Read {len(headers)} columns, {len(data)} rows (no artificial limits)")
        return headers, data

    def _read_data_chunked(self, sheet: xw.Sheet, start_row: int, end_row: int, last_col: int) -> List[List]:
        """Read data in chunks to handle large files efficiently"""
        data = []
        chunk_size = self.config.chunk_size
        
        r = start_row
        chunk_count = 0
        while r <= end_row:
            # Check memory pressure before each chunk
            if MemoryOptimizer.check_memory_pressure():
                print("   ⚠️ Memory pressure detected, reducing chunk size")
                chunk_size = max(chunk_size // 2, 500)
                MemoryOptimizer.cleanup_memory()
            
            r2 = min(r + chunk_size - 1, end_row)
            chunk_count += 1
            
            try:
                print(f"   📦 Reading chunk {chunk_count}: rows {r}-{r2}")
                chunk = EnhancedExcelOptimizer.safe_excel_operation(
                    lambda rr=r, rr2=r2: sheet.range((rr, 1), (rr2, last_col)).value
                )
                
                if chunk:
                    if not isinstance(chunk, list):
                        chunk = [chunk]
                    elif len(chunk) > 0 and not isinstance(chunk[0], list):
                        chunk = [chunk]
                    data.extend(chunk)
                    
            except Exception as e:
                print(f"   ❌ Chunk {chunk_count} failed: {e}")
                # Try smaller chunk on failure
                if chunk_size > 100:
                    chunk_size = chunk_size // 2
                    continue
                else:
                    break
                    
            r = r2 + 1
            
        print(f"   ✅ Completed chunked reading: {chunk_count} chunks, {len(data)} total rows")
        return data

    # ---------- business logic ----------
    def _process_dataframe_enhanced(
        self, df: pd.DataFrame, sheet: xw.Sheet, header_row: int,
        headers: List[str], summary_subset: pd.DataFrame
    ) -> Tuple[int, int]:

        df['Item2'] = df['Item2'].astype(str).str.strip()
        df['Note']  = df['Note'].astype(str).str.strip()
        mask = (df['Item2'] == 'Leasing period') & (df['Note'] == 'Committed')
        df_block = df[mask].copy().reset_index(drop=True)
        print(f"   ✔️ {len(df_block)} existing 'Leasing period' + 'Committed' rows found.")
        if df_block.empty:
            return 0, 0

        df_block['key1*'] = (df_block['Factory code'].astype(str).str.strip() + '|' +
                             df_block['Tenant code'].astype(str).str.strip())
        df_block['key2*'] = (df_block['Factory code'].astype(str).str.strip() + '|' +
                             df_block['Tenant name'].astype(str).str.strip())

        original_indices = df[mask].index.tolist()
        updated_summary_indices = set()
        write_pairs = []

        # update các dòng khớp
        summary_key1 = summary_subset['Unit name'].astype(str).str.strip() + '|' + summary_subset['Tenant ID'].astype(str).str.strip()
        summary_key2 = summary_subset['Unit name'].astype(str).str.strip() + '|' + summary_subset['Tenant'].astype(str).str.strip()

        for i, (_, row) in enumerate(df_block.iterrows()):
            k1, k2 = row['key1*'], row['key2*']
            match = summary_subset[(summary_key1 == str(k1)) | (summary_key2 == str(k2))]
            if not match.empty:
                srow = match.iloc[0]
                updated_summary_indices.add(match.index[0])
                new_vals = []
                for col_name in headers:
                    val = row.get(col_name, '')
                    if col_name in self.config.column_mapping.values():
                        src_col = next((src for src, tgt in self.config.column_mapping.items() if tgt == col_name), None)
                        if src_col in srow.index:
                            cand = srow[src_col]
                            if pd.notna(cand) and str(cand).strip() not in ['', '- None -']:
                                val = cand
                    val = self._ensure_scalar(val)
                    new_vals.append(val)
                excel_row = header_row + 1 + original_indices[i]
                write_pairs.append((excel_row, new_vals))

        rows_updated = 0
        if write_pairs:
            # Optimized: Batch write all updates at once
            try:
                write_pairs.sort(key=lambda x: x[0])
                print(f"   ⚡ Batch updating {len(write_pairs)} rows...")
                
                # Group consecutive rows for range-based updates
                batch_groups = []
                current_group = []
                for excel_row, vals in write_pairs:
                    if not current_group or excel_row == current_group[-1][0] + 1:
                        current_group.append((excel_row, vals))
                    else:
                        batch_groups.append(current_group)
                        current_group = [(excel_row, vals)]
                if current_group:
                    batch_groups.append(current_group)
                
                for group in batch_groups:
                    if len(group) == 1:
                        # Single row update
                        excel_row, vals = group[0]
                        sheet.range((excel_row, 1), (excel_row, len(headers))).value = vals
                        rows_updated += 1
                    else:
                        # Multi-row batch update
                        start_row = group[0][0]
                        end_row = group[-1][0]
                        batch_data = [vals for _, vals in group]
                        sheet.range((start_row, 1), (end_row, len(headers))).value = batch_data
                        rows_updated += len(group)
                        
                print(f"   → Updated {rows_updated} existing rows with summary data")
            except Exception as e:
                print(f"   ⚠️ Batch update failed, falling back to row-by-row: {e}")
                # Fallback to original method
                for excel_row, vals in write_pairs:
                    try:
                        sheet.range((excel_row, 1), (excel_row, len(headers))).value = vals
                        rows_updated += 1
                    except Exception as e:
                        print(f"     ⚠️ Update row {excel_row}: {e}")
                print(f"   → Updated {rows_updated} existing rows with summary data")
        else:
            print("   → No existing rows matched for update.")

        # fill các dòng “green” trống còn lại bằng summary chưa dùng
        # Enhanced auto-fill logic with dynamic row creation
        unmatched_summary = summary_subset.loc[~summary_subset.index.isin(updated_summary_indices)]
        rows_added = 0
        
        if not unmatched_summary.empty:
            # Find existing empty green rows
            empty_green_mask = (
                (df['Item2'].astype(str).str.strip() == 'Leasing period') &
                (df['Note'].astype(str).str.strip() == 'Committed') &
                ((df['Factory code'].astype(str).str.strip() == '') |
                 (df['Tenant code'].astype(str).str.strip() == '') |
                 (df['Tenant name'].astype(str).str.strip() == ''))
            )
            empty_green_rows = df[empty_green_mask]
            
            unmatched_count = len(unmatched_summary)
            empty_count = len(empty_green_rows)
            
            print(f"   → Empty green rows: {empty_count} | Unmatched summary: {unmatched_count}")
            
            # Auto-add more empty green rows if needed
            if empty_count < unmatched_count:
                rows_to_add = unmatched_count - empty_count
                print(f"   🔄 Auto-adding {rows_to_add} empty green rows to match unmatched summary")
                
                # Find the last row in the sheet to add new rows
                try:
                    last_used_row = sheet.used_range.last_cell.row
                    rows_added += self._add_empty_green_rows(sheet, last_used_row + 1, rows_to_add, headers)
                    
                    # Update the dataframe to include new rows for processing
                    for i in range(rows_to_add):
                        new_row_data = [''] * len(headers)
                        # Set the required green row identifiers
                        if 'Item2' in headers:
                            new_row_data[headers.index('Item2')] = 'Leasing period'
                        if 'Note' in headers:
                            new_row_data[headers.index('Note')] = 'Committed'
                        df.loc[len(df)] = new_row_data
                    
                    # Recalculate empty green rows with the new rows
                    empty_green_mask = (
                        (df['Item2'].astype(str).str.strip() == 'Leasing period') &
                        (df['Note'].astype(str).str.strip() == 'Committed') &
                        ((df['Factory code'].astype(str).str.strip() == '') |
                         (df['Tenant code'].astype(str).str.strip() == '') |
                         (df['Tenant name'].astype(str).str.strip() == ''))
                    )
                    empty_green_rows = df[empty_green_mask]
                    print(f"   ✅ Updated empty green rows: {len(empty_green_rows)}")
                    
                except Exception as e:
                    print(f"   ⚠️ Could not auto-add green rows: {e}")
                    # Continue with existing empty rows

            if len(empty_green_rows) > 0:
                empty_excel_rows = [header_row + 1 + idx for idx in empty_green_rows.index.tolist()]
                fill_pairs = []
                for i, (_, srow) in enumerate(unmatched_summary.iterrows()):
                    if i >= len(empty_excel_rows): break
                    excel_row = empty_excel_rows[i]
                    new_vals = []
                    for col_name in headers:
                        val = ''
                        if col_name in self.config.column_mapping.values():
                            src_col = next((src for src, tgt in self.config.column_mapping.items() if tgt == col_name), None)
                            if src_col in srow.index:
                                cand = srow[src_col]
                                if pd.notna(cand) and str(cand).strip() not in ['', '- None -']:
                                    val = cand
                        elif col_name == 'Item2':
                            val = 'Leasing period'
                        elif col_name == 'Note':
                            val = 'Committed'
                        elif col_name == 'Factory code':
                            val = srow.get('Unit name', '')
                        elif col_name == 'Tenant code':
                            val = srow.get('Tenant ID', '')
                        elif col_name == 'Tenant name':
                            val = srow.get('Tenant', '')
                        else:
                            row_idx = empty_green_rows.index[i]
                            current_val = df.iloc[row_idx].get(col_name, '')
                            val = self._ensure_scalar(current_val) if current_val != '' else ''
                        new_vals.append(self._ensure_scalar(val))
                    fill_pairs.append((excel_row, new_vals))

                # Optimized: Batch write all fills at once
                try:
                    print(f"   ⚡ Batch filling {len(fill_pairs)} rows...")
                    
                    # Group consecutive rows for range-based fills
                    fill_groups = []
                    current_group = []
                    fill_pairs.sort(key=lambda x: x[0])
                    
                    for excel_row, vals in fill_pairs:
                        if not current_group or excel_row == current_group[-1][0] + 1:
                            current_group.append((excel_row, vals))
                        else:
                            fill_groups.append(current_group)
                            current_group = [(excel_row, vals)]
                    if current_group:
                        fill_groups.append(current_group)
                    
                    for group in fill_groups:
                        if len(group) == 1:
                            # Single row fill
                            excel_row, vals = group[0]
                            sheet.range((excel_row, 1), (excel_row, len(headers))).value = vals
                            rows_added += 1
                        else:
                            # Multi-row batch fill
                            start_row = group[0][0]
                            end_row = group[-1][0]
                            batch_data = [vals for _, vals in group]
                            sheet.range((start_row, 1), (end_row, len(headers))).value = batch_data
                            rows_added += len(group)
                            
                except Exception as e:
                    print(f"   ⚠️ Batch fill failed, falling back to row-by-row: {e}")
                    # Fallback to original method
                    for excel_row, vals in fill_pairs:
                        try:
                            sheet.range((excel_row, 1), (excel_row, len(headers))).value = vals
                            rows_added += 1
                        except Exception as e:
                            print(f"     ⚠️ Fill row {excel_row}: {e}")
                            
                print(f"   → Filled {rows_added} empty green rows")
            else:
                print("   ⚠️ No empty green rows to fill")
        else:
            print("   → No unmatched summary rows to fill")

        return rows_updated, rows_added

    @staticmethod
    def _ensure_scalar(val):
        if hasattr(val, 'iloc') and len(val) > 0:
            return val.iloc[0]
        if hasattr(val, 'item'):
            return val.item()
        if pd.isna(val):
            return ''
        return val

    def _add_empty_green_rows(self, sheet: xw.Sheet, start_row: int, count: int, headers: List[str]) -> int:
        """Add empty green rows to the sheet for auto-fill functionality"""
        try:
            rows_added = 0
            for i in range(count):
                row_num = start_row + i
                
                # Create empty row data with proper identifiers
                row_data = [''] * len(headers)
                
                # Set the required green row identifiers
                if 'Item2' in headers:
                    row_data[headers.index('Item2')] = 'Leasing period'
                if 'Note' in headers:
                    row_data[headers.index('Note')] = 'Committed'
                
                # Write the row to Excel
                sheet.range((row_num, 1), (row_num, len(headers))).value = row_data
                rows_added += 1
                
            print(f"   ➕ Added {rows_added} empty green rows at row {start_row}")
            return rows_added
            
        except Exception as e:
            print(f"   ❌ Failed to add empty green rows: {e}")
            return 0
