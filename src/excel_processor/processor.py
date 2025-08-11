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
        
        # First, validate file exists and basic format
        import os
        if not os.path.exists(summary_path):
            raise FileNotFoundError(f"Summary file not found: {summary_path}")
        
        file_size = os.path.getsize(summary_path)
        if file_size == 0:
            raise ValueError(f"Summary file is empty: {summary_path}")
        
        print(f"   📁 File: {os.path.basename(summary_path)} ({file_size:,} bytes)")
        
        # Check file signature to detect actual format
        def detect_file_format(filepath):
            try:
                with open(filepath, 'rb') as f:
                    header = f.read(512)  # Read more bytes for better detection
                    
                if header.startswith(b'PK\x03\x04'):
                    return 'xlsx'  # ZIP-based format (xlsx, xlsm)
                elif header.startswith(b'\xd0\xcf\x11\xe0'):
                    return 'xls'   # OLE2/CFB format (xls)
                elif header.startswith(b'<?xml'):
                    # Check if it's Excel XML format
                    header_str = header.decode('utf-8', errors='ignore')
                    if 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' in header_str:
                        return 'excel_xml'  # Excel XML Spreadsheet
                    else:
                        return 'xml'   # Generic XML format
                else:
                    return 'unknown'
            except:
                return 'unknown'
        
        actual_format = detect_file_format(summary_path)
        file_ext = os.path.splitext(summary_path.lower())[1]
        
        print(f"   🔍 File extension: {file_ext}, Detected format: {actual_format}")
        
        # Handle Excel XML format specially
        if actual_format == 'excel_xml':
            try:
                print("   🔧 Detected Excel XML format - using custom XML parser")
                import xml.etree.ElementTree as ET
                
                # Parse Excel XML format
                tree = ET.parse(summary_path)
                root = tree.getroot()
                
                # Find worksheets and data
                ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
                worksheets = root.findall('.//ss:Worksheet', ns)
                
                if not worksheets:
                    raise ValueError("No worksheets found in Excel XML file")
                
                # Use first worksheet (or find by name if specified)
                worksheet = worksheets[0]
                worksheet_name = worksheet.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', 'Sheet1')
                print(f"   📋 Using worksheet: {worksheet_name}")
                
                # Get table data
                table = worksheet.find('.//ss:Table', ns)
                if table is None:
                    raise ValueError("No table found in worksheet")
                
                rows = table.findall('.//ss:Row', ns)
                
                # Extract data from XML with better cell handling
                data_rows = []
                headers = None
                
                for row_idx, row_elem in enumerate(rows):
                    cells = row_elem.findall('.//ss:Cell', ns)
                    row_data = []
                    current_col = 0
                    
                    for cell in cells:
                        # Handle cell index (Excel XML can have sparse cells)
                        cell_index = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                        if cell_index:
                            target_col = int(cell_index) - 1  # Convert to 0-based
                            # Fill gaps with empty strings
                            while current_col < target_col:
                                row_data.append('')
                                current_col += 1
                        
                        # Extract cell value
                        data_elem = cell.find('.//ss:Data', ns)
                        if data_elem is not None and data_elem.text:
                            cell_value = str(data_elem.text).strip()
                        else:
                            cell_value = ''
                        
                        row_data.append(cell_value)
                        current_col += 1
                    
                    # Skip completely empty rows
                    if any(cell.strip() for cell in row_data if cell):
                        if headers is None and row_data:
                            headers = row_data
                            print(f"   📋 Found headers: {len(headers)} columns")
                            # Show first few column names for debugging
                            preview_cols = headers[:5] + (['...'] if len(headers) > 5 else [])
                            print(f"   🔍 Column preview: {preview_cols}")
                        elif headers is not None:
                            # Ensure row has same length as headers
                            while len(row_data) < len(headers):
                                row_data.append('')
                            data_rows.append(row_data[:len(headers)])  # Trim if too long
                
                if not headers:
                    raise ValueError("No headers found in Excel XML file")
                
                if not data_rows:
                    raise ValueError("No data rows found in Excel XML file")
                
                # Create DataFrame
                self.summary_data = pd.DataFrame(data_rows, columns=headers).fillna('').astype(str)
                print(f"   ✅ Successfully loaded Excel XML: {len(data_rows)} rows × {len(headers)} columns")
                
                # Debug: Print actual column names to help identify the issue
                print(f"   🔍 Loaded columns: {list(self.summary_data.columns)}")
                
                # Check for 'Subsidiary' column with different variations
                subsidiary_cols = [col for col in self.summary_data.columns if 'subsidiary' in col.lower()]
                if subsidiary_cols:
                    print(f"   ✅ Found subsidiary column(s): {subsidiary_cols}")
                else:
                    print(f"   ⚠️  No 'Subsidiary' column found. Available columns: {list(self.summary_data.columns)[:10]}")
                
            except Exception as e:
                print(f"   ❌ Excel XML parsing failed: {str(e)}")
                # Fallback to trying pandas engines
                actual_format = 'xml'
                if hasattr(self, 'summary_data'):
                    delattr(self, 'summary_data')
        
        # For non-Excel XML formats, use pandas engines
        if actual_format != 'excel_xml' or not hasattr(self, 'summary_data') or self.summary_data is None:
            # Determine engines to try based on detected format
            engines_to_try = []
            
            if actual_format == 'xlsx':
                engines_to_try = ['openpyxl']
            elif actual_format == 'xls':
                engines_to_try = ['xlrd']
            elif actual_format == 'xml':
                # For XML files, try engines that might handle XML
                engines_to_try = ['openpyxl', 'xlrd']
            else:
                # Unknown format - try all engines
                if file_ext == '.xls':
                    engines_to_try = ['xlrd', 'openpyxl']
                elif file_ext == '.xlsx':
                    engines_to_try = ['openpyxl']
                elif file_ext == '.xlsb':
                    engines_to_try = ['pyxlsb', 'openpyxl']
                else:
                    engines_to_try = ['openpyxl', 'xlrd', 'pyxlsb']
            
            # Try each engine until one works
            last_error = None
            successful_engine = None
            
            for engine in engines_to_try:
                try:
                    print(f"   🔧 Trying engine: {engine}")
                    self.summary_data = pd.read_excel(summary_path, dtype=str, engine=engine).fillna('')
                    successful_engine = engine
                    print(f"   ✅ Successfully loaded with {engine} engine")
                    print(f"   🔍 Pandas loaded columns: {list(self.summary_data.columns)[:10]}")
                    break
                except Exception as e:
                    error_msg = str(e)
                    print(f"   ❌ {engine} failed: {error_msg[:100]}...")
                    last_error = e
                    continue
            
            if successful_engine is None and not hasattr(self, 'summary_data'):
                # Provide detailed error message with file analysis
                error_details = [
                    f"\n❌ SUMMARY FILE LOADING FAILED",
                    f"File: {summary_path}",
                    f"File size: {file_size:,} bytes",
                    f"Extension: {file_ext}",
                    f"Detected format: {actual_format}",
                    f"Engines tried: {', '.join(engines_to_try)}",
                    f"Last error: {str(last_error)}"
                ]
                
                # Provide helpful suggestions
                suggestions = [
                    "\n💡 POSSIBLE SOLUTIONS:",
                    "1. Check if the file is corrupted or incomplete",
                    "2. Try opening the file in Excel to verify it's valid",
                    "3. Resave the file as a proper Excel format (.xlsx recommended)",
                    "4. For XML Excel files, ensure proper XML structure",
                    "5. Try converting to .xlsx format in Excel: File > Save As > Excel Workbook",
                    "6. Check file permissions and access rights"
                ]
                
                if "zip file" in str(last_error).lower():
                    suggestions.append("7. File appears corrupted - the Excel file structure is damaged")
                
                if actual_format == 'xml':
                    suggestions.append("8. XML format detected - try opening in Excel and saving as .xlsx")
                
                full_error = "\n".join(error_details + suggestions)
                raise Exception(full_error)

        # Validate loaded data structure
        if self.summary_data.empty:
            raise ValueError(f"Summary file loaded successfully but contains no data")
        
        print(f"   📊 Final data shape: {self.summary_data.shape[0]} rows × {self.summary_data.shape[1]} columns")
        
        # Check for required columns with flexible matching
        required_columns = ['Subsidiary', 'Unit name', 'Tenant ID', 'Tenant']
        available_columns = list(self.summary_data.columns)
        
        # Try to find columns with flexible matching
        column_mapping = {}
        for req_col in required_columns:
            # Exact match first
            if req_col in available_columns:
                column_mapping[req_col] = req_col
                continue
            
            # Case-insensitive match
            matches = [col for col in available_columns if col.lower() == req_col.lower()]
            if matches:
                column_mapping[req_col] = matches[0]
                continue
                
            # Partial match (contains the required column name)
            matches = [col for col in available_columns if req_col.lower() in col.lower()]
            if matches:
                column_mapping[req_col] = matches[0]
                print(f"   🔄 Mapped '{req_col}' to '{matches[0]}'")
                continue
        
        missing_columns = [col for col in required_columns if col not in column_mapping]
        if missing_columns:
            print(f"   ⚠️  Missing expected columns: {missing_columns}")
            print(f"   📋 Available columns: {available_columns[:20]}")
            # Don't fail, just warn - the process might still work
        
        # Use mapped column names for processing
        if 'Subsidiary' in column_mapping:
            subsidiary_col = column_mapping['Subsidiary']
            subsidiaries = self.summary_data[subsidiary_col].unique()
        else:
            print("   ⚠️  No Subsidiary column found - using first column as subsidiary")
            subsidiary_col = available_columns[0] if available_columns else 'Column1'
            subsidiaries = self.summary_data[subsidiary_col].unique() if subsidiary_col in self.summary_data.columns else []
        
        for sub in subsidiaries:
            if pd.notna(sub) and str(sub).strip():
                clean = str(sub).strip().upper()
                self.subsidiary_variations[clean] = str(sub)
                if '-' in clean:
                    self.subsidiary_variations[clean.split('-')[0].strip()] = str(sub)

        # Build lookup with mapped column names
        unit_name_col = column_mapping.get('Unit name', 'Unit name')
        tenant_id_col = column_mapping.get('Tenant ID', 'Tenant ID')
        tenant_col = column_mapping.get('Tenant', 'Tenant')
        
        self.summary_lookup = {}
        for idx, row in self.summary_data.iterrows():
            unit_name = str(row.get(unit_name_col, '')).strip()
            tenant_id = str(row.get(tenant_id_col, '')).strip()
            tenant = str(row.get(tenant_col, '')).strip()
            
            k1 = f"{unit_name}|{tenant_id}"
            k2 = f"{unit_name}|{tenant}"
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
            
            # Auto-add more empty green rows if needed (with double buffer for future growth)
            if empty_count < unmatched_count:
                # Calculate rows needed with buffer
                base_rows_needed = unmatched_count - empty_count
                buffer_rows = max(base_rows_needed, 10)  # Double buffer, minimum 10 rows
                total_rows_to_add = base_rows_needed + buffer_rows
                
                print(f"   🔄 Auto-adding {total_rows_to_add} green rows ({base_rows_needed} needed + {buffer_rows} buffer)")
                
                # Find the best insertion point: right after the last existing green row
                try:
                    # Find all existing green rows to determine insertion point
                    green_mask = (
                        (df['Item2'].astype(str).str.strip() == 'Leasing period') &
                        (df['Note'].astype(str).str.strip() == 'Committed')
                    )
                    green_indices = df[green_mask].index.tolist()
                    
                    if green_indices:
                        # Insert after the last green row
                        last_green_df_idx = max(green_indices)
                        insertion_excel_row = header_row + 1 + last_green_df_idx
                        
                        # Find a good reference row for formatting (use the last green row)
                        reference_excel_row = insertion_excel_row
                        
                        print(f"   📍 Inserting {total_rows_to_add} rows after last green row (Excel row {insertion_excel_row})")
                        
                        # Use the new formatted row addition method
                        added_count = self._add_formatted_green_rows(
                            sheet, reference_excel_row, insertion_excel_row, total_rows_to_add, headers
                        )
                        rows_added += added_count
                        
                        print(f"   ✅ Added {added_count} formatted rows with preserved styling")
                    else:
                        # Fallback: add at end if no green rows found
                        print("   ⚠️ No existing green rows found for reference, adding at end")
                        last_used_row = sheet.used_range.last_cell.row
                        added_count = self._add_empty_green_rows(sheet, last_used_row + 1, total_rows_to_add, headers)
                        rows_added += added_count
                    
                    # Re-read the sheet data to include newly inserted rows
                    print("   📊 Re-reading sheet data to include new formatted rows...")
                    try:
                        # Re-read the data from Excel to get the updated structure
                        updated_headers, updated_data = self._batch_read_enhanced(sheet, header_row)
                        
                        # Reconstruct the DataFrame with the new data
                        df = pd.DataFrame(updated_data, columns=updated_headers).astype(object).fillna('')
                        print(f"   ✅ Updated DataFrame: {len(df)} rows (including new green rows)")
                        
                    except Exception as e:
                        print(f"   ⚠️ Failed to re-read sheet, manually adding rows to DataFrame: {e}")
                        # Fallback: manually add rows to the existing DataFrame
                        for i in range(total_rows_to_add):
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

    def _add_formatted_green_rows(self, sheet: xw.Sheet, reference_row: int, insert_after_row: int, count: int, headers: List[str]) -> int:
        """
        Add empty green rows with complete formatting preservation.
        
        Args:
            sheet: Excel worksheet
            reference_row: Row number to copy formatting from (existing green row)
            insert_after_row: Row number after which to insert new rows
            count: Number of rows to add
            headers: Column headers list
        
        Returns:
            Number of rows successfully added
        """
        try:
            if count <= 0:
                return 0
                
            print(f"   🎨 Adding {count} formatted green rows after row {insert_after_row}")
            print(f"   📋 Using row {reference_row} as formatting template")
            
            rows_added = 0
            
            # Insert rows first to make space (this preserves relative positioning)
            insertion_point = insert_after_row + 1
            
            # Insert multiple rows at once for better performance
            try:
                # Insert rows by selecting a range and using Excel's insert functionality
                insert_range = sheet.range((insertion_point, 1), (insertion_point + count - 1, len(headers)))
                insert_range.api.EntireRow.Insert()
                print(f"   ✅ Inserted {count} blank rows at position {insertion_point}")
            except Exception as e:
                print(f"   ⚠️ Bulk row insertion failed, trying one-by-one: {e}")
                # Fallback: insert rows one by one
                for i in range(count):
                    try:
                        sheet.range((insertion_point, 1), (insertion_point, len(headers))).api.EntireRow.Insert()
                    except Exception as row_error:
                        print(f"   ❌ Failed to insert row {i+1}: {row_error}")
                        break
            
            # Now copy formatting and data from the reference row
            for i in range(count):
                target_row = insertion_point + i
                
                try:
                    # Copy entire row formatting from reference row
                    source_range = sheet.range((reference_row, 1), (reference_row, len(headers)))
                    target_range = sheet.range((target_row, 1), (target_row, len(headers)))
                    
                    # Copy formatting (background color, borders, fonts, number formats)
                    source_range.api.Copy()
                    target_range.api.PasteSpecial(-4122)  # xlPasteFormats
                    
                    # Copy formulas if any exist
                    try:
                        source_range.api.Copy()
                        target_range.api.PasteSpecial(-4123)  # xlPasteFormulas
                    except:
                        pass  # No formulas to copy
                    
                    # Clear clipboard
                    try:
                        sheet.app.api.CutCopyMode = False
                    except:
                        pass
                    
                    # Set the data values for green row identification
                    row_data = [''] * len(headers)
                    
                    # Set required identifiers
                    if 'Item2' in headers:
                        row_data[headers.index('Item2')] = 'Leasing period'
                    if 'Note' in headers:
                        row_data[headers.index('Note')] = 'Committed'
                    
                    # Write only the data, preserving all formatting
                    target_range.value = row_data
                    rows_added += 1
                    
                except Exception as e:
                    print(f"   ⚠️ Failed to format row {target_row}: {e}")
                    # Still try to add basic data even if formatting fails
                    try:
                        row_data = [''] * len(headers)
                        if 'Item2' in headers:
                            row_data[headers.index('Item2')] = 'Leasing period'
                        if 'Note' in headers:
                            row_data[headers.index('Note')] = 'Committed'
                        sheet.range((target_row, 1), (target_row, len(headers))).value = row_data
                        rows_added += 1
                    except Exception as data_error:
                        print(f"   ❌ Failed to add data to row {target_row}: {data_error}")
                        break
            
            print(f"   ✅ Successfully added {rows_added} formatted green rows")
            return rows_added
            
        except Exception as e:
            print(f"   ❌ Failed to add formatted green rows: {e}")
            return 0

    def _add_empty_green_rows(self, sheet: xw.Sheet, start_row: int, count: int, headers: List[str]) -> int:
        """Legacy method - kept for backward compatibility"""
        print(f"   ⚠️ Using legacy row addition method - formatting may not be preserved")
        return self._add_formatted_green_rows(sheet, start_row - 1, start_row - 1, count, headers)
