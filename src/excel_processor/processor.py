# excel_processor/processor.py
import time
import re
from datetime import datetime, timedelta, timezone
import pandas as pd
import xlwings as xw
from typing import List, Tuple, Optional, Dict

from .models import ProcessingConfig, ProcessingResult
from .com_management import COMManager, EnhancedExcelOptimizer
from .subsidiary import SubsidiaryExtractor
from .memory_optimizer import MemoryOptimizer
from .summary_comparator import SummaryComparator

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

        # Logging optimization
        self.verbose_logging = getattr(config, 'verbose_logging', False)
        self._date_parse_cache = {}  # Cache for date parsing to avoid redundant parsing

        # Feature flags for optimization toggles
        self.feature_flags = {
            'enable_date_filtering': getattr(config, 'enable_date_filtering', True),
            'enable_caching': getattr(config, 'enable_caching', True),
            'enable_row_highlighting': getattr(config, 'enable_row_highlighting', True),
            'enable_batch_updates': getattr(config, 'enable_batch_updates', True),
            'enable_project_code_parsing': getattr(config, 'enable_project_code_parsing', True),
            'enable_auto_add_rows': getattr(config, 'enable_auto_add_rows', True)
        }
    def _parse_date_flexible(self, date_str: str) -> Optional[datetime]:
        """Parse date from multiple formats commonly found in Excel (with caching)"""
        if not date_str or pd.isna(date_str):
            return None
            
        date_str = str(date_str).strip()
        if not date_str or date_str.lower() in ['', 'nan', 'none', 'null']:
            return None
        
        # Check cache first
        if self._is_feature_enabled('enable_caching') and date_str in self._date_parse_cache:
            return self._date_parse_cache[date_str]
                
        # For summary data, prioritize mm/dd/yyyy format but also support other common formats
        date_formats = [
            "%m/%d/%Y",              # mm/dd/yyyy (US format: month/day/year) - PRIORITY for summary
            "%m/%d/%y",              # mm/dd/yy (US format with 2-digit year)
            "%Y-%m-%d %H:%M:%S",     # 2022-10-30 00:00:00 (datetime format from summary)
            "%Y-%m-%d",              # 2022-10-30 (ISO date format)
            "%d-%b-%y",              # 8-Mar-24
            "%d-%B-%y",              # 8-March-24
            "%d-%b-%Y",              # 8-Mar-2024
            "%d-%B-%Y",              # 8-March-2024
            "%Y/%m/%d",              # 2023/03/08
        ]
        
        # Try each format in order - mm/dd/yyyy will be tried first
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        try:
            # Force US locale interpretation for mm/dd/yyyy
            parsed = pd.to_datetime(date_str, errors='coerce')
            if pd.notna(parsed):
               return parsed.to_pydatetime()
        except:
            pass
        
        # Handle Excel serial dates
        try:
            float_val = float(date_str)
            if 1 <= float_val <= 100000:
                base_date = datetime(1899, 12, 30)
                return base_date + timedelta(days=float_val)
        except:
            pass
        
        return None

    def _format_date_consistent(self, date_str: str) -> str:
        """Format date string consistently as MM/DD/YYYY"""
        if not date_str or pd.isna(date_str):
            return ''
        
        parsed_date = self._parse_date_flexible(str(date_str))
        if parsed_date:
            # Always format as MM/DD/YYYY (zero-padded)
            return parsed_date.strftime('%m/%d/%Y')
        else:
            # If parsing fails, return original string
            return str(date_str).strip()

    def _is_within_90_days(self, date_str: str) -> bool:
        """Check if the given date is 90 days before today or any day after today"""
        parsed_date = self._parse_date_flexible(date_str)
        if parsed_date is None:
            return False

        today = datetime.now()
        lower_bound = today - timedelta(days=90)

        return parsed_date >= lower_bound

    def _parse_phase_column(self, phase_str: str) -> dict:

        if not phase_str or pd.isna(phase_str):
            return {"project_code": None, "project_name": None, "phase": None}
        
        phase_str = str(phase_str).strip()
        if not phase_str:
            return {"project_code": None, "project_name": None, "phase": None}
              
        # Direct processing without cache lookup
        parts = phase_str.split(':', 1)
        if len(parts) < 2:
            # No colon found, treat entire string as project code
            return {"project_code": phase_str, "project_name": None, "phase": None}
        project_code = parts[0].strip()
        description = parts[1].strip()
            
        # Split description to separate project name from phase
        # Look for "_Phase" pattern
        if "_Phase" in description:
            desc_parts = description.rsplit("_Phase", 1)
            project_name = desc_parts[0].strip()
            phase = f"Phase{desc_parts[1]}".strip() if len(desc_parts) > 1 else None
        else:
            project_name = description
            phase = None
        
        return {
            "project_code": project_code,
            "project_name": project_name,
            "phase": phase
        }

    def _apply_row_formatting(self, sheet, excel_row: int, headers_count: int, format_type: str):
        """Apply color formatting to a row
        
        format_type: 'update' for red background, 'add' for yellow background
        """
        try:
            range_obj = sheet.range((excel_row, 1), (excel_row, headers_count))
            if format_type == 'update':
                range_obj.color = (255, 200, 200)  # Light red for updated rows
            elif format_type == 'add':
                range_obj.color = (255, 255, 200)  # Light yellow for added rows
        except Exception as e:
            print(f"   ⚠️ Could not apply formatting to row {excel_row}: {e}")

    def _log_info(self, message: str):
        """Always log important information"""
        print(message)

    def _is_feature_enabled(self, feature_name: str) -> bool:
        """Check if a feature flag is enabled"""
        return self.feature_flags.get(feature_name, True)

    def _optimize_dataframe_memory(self, df: pd.DataFrame) -> pd.DataFrame:
        """Optimize DataFrame memory usage by downcasting numeric types"""
        if not self._is_feature_enabled('enable_memory_optimization'):
            return df
            
        original_memory = df.memory_usage(deep=True).sum()
        
        # Convert object columns to category if they have few unique values
        for col in df.select_dtypes(include=['object']):
            if df[col].nunique() / len(df) < 0.5:  # Less than 50% unique values
                df[col] = df[col].astype('category')
        
        # Downcast numeric columns
        for col in df.select_dtypes(include=['int']):
            df[col] = pd.to_numeric(df[col], downcast='integer')
        
        for col in df.select_dtypes(include=['float']):
            df[col] = pd.to_numeric(df[col], downcast='float')
        
        new_memory = df.memory_usage(deep=True).sum()
        memory_saved = original_memory - new_memory
        
        if memory_saved > 0:
            self._log_verbose(f"   📊 Memory optimized: {memory_saved / 1024 / 1024:.2f}MB saved")
        
        return df

    def _get_optimal_batch_size(self, total_rows: int) -> int:
        """Calculate optimal batch size based on data size and memory"""
        if total_rows < 1000:
            return min(total_rows, 100)
        elif total_rows < 10000:
            return 200
        else:
            return 500
    
    def _validate_data_integrity(self, df: pd.DataFrame, required_columns: list) -> bool:
        """Quick validation of data integrity"""
        if df.empty:
            return False
        
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            self._log_info(f"   ⚠️ Missing required columns: {missing_cols}")
            return False
        
        return True

    def _filter_summary_by_start_date(self, summary_subset: pd.DataFrame) -> pd.DataFrame:
        """Filter summary data to only include records with Start date within 90 days"""
        if summary_subset.empty:
            return summary_subset
        
        # Look for Start date column
        start_date_columns = []
        for col in summary_subset.columns:
            col_lower = col.lower().strip()
            if any(keyword in col_lower for keyword in ['start', 'begin', 'commence', 'effective']):
                if any(keyword in col_lower for keyword in ['date', 'time', 'day']):
                    start_date_columns.append(col)
        
        # Also check for exact matches
        exact_matches = [col for col in summary_subset.columns if col.lower().strip() == 'start date']
        start_date_columns.extend(exact_matches)
        
        # Remove duplicates
        start_date_columns = list(dict.fromkeys(start_date_columns))
        
        if not start_date_columns:
            print("   ⚠️ No Start date column found - skipping 90-day filter")
            return summary_subset
        
        # Use the first matching column
        start_date_col = start_date_columns[0]
        print(f"   📅 Using '{start_date_col}' column for 90-day filtering")
        
        # Apply the filter
        original_count = len(summary_subset)
        within_90_days_mask = summary_subset[start_date_col].apply(self._is_within_90_days)
        filtered_subset = summary_subset[within_90_days_mask].copy()
        
        filtered_count = len(filtered_subset)
        excluded_count = original_count - filtered_count
        
        print(f"   📊 90-day filter: {filtered_count} included, {excluded_count} excluded")
        
        return filtered_subset

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
        
        # FALLBACK: If no subsidiary match found, return all data with warning
        print(f"   ⚠️ No subsidiary match for '{extracted_subsidiary}' - processing ALL summary data")
        result = self.summary_data if not hasattr(self, '_already_filtered') else self.summary_data
        # Ensure Start date column is consistently formatted as MM/DD/YYYY
        if 'Start date' in result.columns:
            result['Start date'] = result['Start date'].apply(self._format_date_consistent)
        
        return self._filter_summary_by_start_date(result) # Process everything instead of empty DataFrame


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

    # Fixed version of _process_dataframe_enhanced method
    def _process_dataframe_enhanced(
        self, df: pd.DataFrame, sheet: xw.Sheet, header_row: int,
        headers: List[str], summary_subset: pd.DataFrame
    ) -> Tuple[int, int]:
        """
        OPTIMIZED: Enhanced processing with vectorized operations and O(n) complexity.
        Filters summary data by 90-day date condition and ensures no duplicates.
        """
        import time
        start_time = time.time()
        
        # OPTIMIZATION 1: Single-pass vectorized preprocessing
        print("   🔄 Preprocessing data...")
        
        # Vectorized string operations - do once, use multiple times
        df_item2_clean = df['Item2'].astype('string').str.strip().str.lower()
        df_note_clean = df['Note'].astype('string').str.strip().str.lower()
        
        # OPTIMIZATION 2: Vectorized mask creation
        existing_mask = (df_item2_clean == 'leasing period') & (df_note_clean == 'committed')
        df_block = df[existing_mask].copy()
        
        print(f"   ✅ {len(df_block)} existing 'Leasing period' + 'Committed' rows found.")
        
        # OPTIMIZATION 3: Efficient date filtering with early exit
        if summary_subset.empty:
            print("   ⚠️ No summary data provided")
            return 0, 0
            
        filtered_summary = summary_subset.copy()
        if 'Start date' in filtered_summary.columns:
            # Vectorized date filtering using list comprehension (faster than apply)
            date_mask = [
                self._is_within_90_days(str(x)) if pd.notna(x) else False 
                for x in filtered_summary['Start date']
            ]
            filtered_summary = filtered_summary[date_mask].copy()
            print(f"   📅 Filtered to {len(filtered_summary)} rows from current year onwards (from {len(summary_subset)} total)")
        
        if filtered_summary.empty:
            print("   ⚠️ No summary rows from current year onwards")
            return 0, 0
        
        # OPTIMIZATION 4: O(n) duplicate detection using vectorized operations
        print("   🔍 Building duplicate detection keys...")
        existing_keys = self._build_existing_keys_vectorized(df_block)
        new_summary_rows = self._filter_duplicates_vectorized(filtered_summary, existing_keys)
        
        print(f"   🔍 Created {len(existing_keys)} existing key combinations")
        print(f"   ➕ {len(new_summary_rows)} new rows to add")
        
        if not new_summary_rows:
            print("   ➡️ No new rows to add (all filtered summary rows already exist)")
            return 0, 0
        
        # OPTIMIZATION 5: Batch processing for row operations
        rows_added = self._process_new_rows_optimized(
            sheet, df, headers, header_row, new_summary_rows, df_item2_clean, df_note_clean
        )
        
        processing_time = time.time() - start_time
        print(f"   ⏱️ Processing completed in {processing_time:.2f}s")
        
        return 0, rows_added  # rows_updated always 0 since we don't update existing rows

    def _build_existing_keys_vectorized(self, df_block: pd.DataFrame) -> set:
        """OPTIMIZED: Build existing keys using vectorized operations - O(n)"""
        if df_block.empty:
            return set()
        
        # Vectorized normalization
        factory_codes = df_block.get('Factory code', pd.Series(dtype='string')).astype('string').str.strip().str.lower()
        tenant_codes = df_block.get('Tenant code', pd.Series(dtype='string')).astype('string').str.strip().str.lower()
        tenant_names = df_block.get('Tenant name', pd.Series(dtype='string')).astype('string').str.strip().str.lower()
        
        existing_keys = set()
        
        # Vectorized key generation using pandas operations
        for idx in df_block.index:
            fc = factory_codes.get(idx, '') or ''
            tc = tenant_codes.get(idx, '') or ''
            tn = tenant_names.get(idx, '') or ''
            
            # Only add non-empty key combinations
            if fc and tc:
                existing_keys.add(f"{fc}|{tc}")
            if fc and tn:
                existing_keys.add(f"{fc}|{tn}")
            if tc and tn:
                existing_keys.add(f"{tc}|{tn}")
        
        return existing_keys

    def _filter_duplicates_vectorized(self, summary_df: pd.DataFrame, existing_keys: set) -> List:
        """OPTIMIZED: Filter duplicates using vectorized operations - O(n)"""
        # Pre-compute normalized columns once
        unit_names = summary_df.get('Unit name', pd.Series(dtype='string')).astype('string').str.strip().str.lower()
        tenant_ids = summary_df.get('Tenant ID', pd.Series(dtype='string')).astype('string').str.strip().str.lower()
        tenants = summary_df.get('Tenant', pd.Series(dtype='string')).astype('string').str.strip().str.lower()
        
        new_rows = []
        
        # Single pass through data
        for idx in summary_df.index:
            un = unit_names.get(idx, '') or ''
            ti = tenant_ids.get(idx, '') or ''
            t = tenants.get(idx, '') or ''
            
            # Generate check keys
            check_keys = []
            if un and ti:
                check_keys.append(f"{un}|{ti}")
            if un and t:
                check_keys.append(f"{un}|{t}")
            if ti and t:
                check_keys.append(f"{ti}|{t}")
            
            # Fast intersection check - short-circuit on first match
            if not any(key in existing_keys for key in check_keys):
                new_rows.append((idx, summary_df.loc[idx]))
        
        return new_rows

    def _process_new_rows_optimized(
        self, sheet: xw.Sheet, df: pd.DataFrame, headers: List[str], 
        header_row: int, new_summary_rows: List, df_item2_clean: pd.Series, df_note_clean: pd.Series
    ) -> int:
        """OPTIMIZED: Process new rows with batch operations and minimal Excel calls"""
        
        # OPTIMIZATION: Reuse pre-computed clean columns
        empty_green_mask = (
            (df_item2_clean == 'leasing period') & 
            (df_note_clean == 'committed') & 
            (df['Factory code'].astype('string').str.strip() == '') &
            (df['Tenant code'].astype('string').str.strip() == '') &
            (df['Tenant name'].astype('string').str.strip() == '')
        )
        empty_green_rows = df[empty_green_mask]
        
        empty_count = len(empty_green_rows)
        new_rows_count = len(new_summary_rows)
        
        print(f"   📊 Empty green rows available: {empty_count}")
        print(f"   📊 New summary rows to add: {new_rows_count}")
        
        # Auto-add green rows if needed
        if empty_count < new_rows_count:
            rows_needed = new_rows_count - empty_count
            buffer_rows = max(5, rows_needed)  # Smaller buffer for efficiency
            total_rows_to_add = rows_needed + buffer_rows
            
            print(f"   🔄 Auto-adding {total_rows_to_add} green rows ({rows_needed} needed + {buffer_rows} buffer)")
            
            if self._auto_add_green_rows(sheet, df, headers, header_row, total_rows_to_add):
                # Re-read and recalculate after adding rows
                updated_headers, updated_data = self._batch_read_enhanced(sheet, header_row)
                df = pd.DataFrame(updated_data, columns=updated_headers).astype(object).fillna('')
                
                # Recalculate with new data
                df_item2_clean = df['Item2'].astype('string').str.strip().str.lower()
                df_note_clean = df['Note'].astype('string').str.strip().str.lower()
                
                empty_green_mask = (
                    (df_item2_clean == 'leasing period') & 
                    (df_note_clean == 'committed') & 
                    (df['Factory code'].astype('string').str.strip() == '') &
                    (df['Tenant code'].astype('string').str.strip() == '') &
                    (df['Tenant name'].astype('string').str.strip() == '')
                )
                empty_green_rows = df[empty_green_mask]
                print(f"   ✅ Updated empty green rows: {len(empty_green_rows)}")
        
        # OPTIMIZATION: Batch fill operations
        return self._batch_fill_rows(sheet, df, headers, header_row, empty_green_rows, new_summary_rows)

    def _auto_add_green_rows(self, sheet: xw.Sheet, df: pd.DataFrame, headers: List[str], 
                           header_row: int, total_rows_to_add: int) -> bool:
        """Add green rows with error handling"""
        try:
            # Find existing green rows for reference
            green_mask = (
                (df['Item2'].astype('string').str.strip() == 'Leasing period') &
                (df['Note'].astype('string').str.strip() == 'Committed')
            )
            green_indices = df[green_mask].index.tolist()
            
            if green_indices:
                last_green_df_idx = max(green_indices)
                insertion_excel_row = header_row + 1 + last_green_df_idx
                
                added_count = self._add_formatted_green_rows(
                    sheet, insertion_excel_row, insertion_excel_row, total_rows_to_add, headers
                )
                return added_count > 0
                
        except Exception as e:
            print(f"   ⚠️ Could not auto-add green rows: {e}")
            
        return False

    def _batch_fill_rows(self, sheet: xw.Sheet, df: pd.DataFrame, headers: List[str],
                        header_row: int, empty_green_rows: pd.DataFrame, new_summary_rows: List) -> int:
        """OPTIMIZED: Batch fill rows with minimal Excel COM calls"""
        
        if empty_green_rows.empty:
            print("   ⚠️ No empty green rows available to fill")
            return 0
        
        empty_excel_rows = [header_row + 1 + idx for idx in empty_green_rows.index.tolist()]
        
        # OPTIMIZATION: Pre-build all row data before any Excel operations
        batch_data = []
        rows_to_process = min(len(empty_excel_rows), len(new_summary_rows))
        
        print(f"   ⚡ Preparing {rows_to_process} rows for batch processing...")
        
        for i in range(rows_to_process):
            excel_row = empty_excel_rows[i]
            summary_idx, srow = new_summary_rows[i]
            
            # Get current row for first column preservation
            current_row_idx = empty_green_rows.index.tolist()[i]
            current_row = df.iloc[current_row_idx]
            
            new_vals = self._build_row_values_optimized(headers, current_row, srow)
            batch_data.append((excel_row, new_vals))
        
        # OPTIMIZATION: Batch write to Excel with error recovery
        return self._batch_write_to_excel(sheet, batch_data, headers)

    def _build_row_values_optimized(self, headers: List[str], current_row: pd.Series, srow: pd.Series) -> List:
        """OPTIMIZED: Build row values with efficient column mapping"""
        new_vals = []
        column_mapping = self.config.column_mapping
        
        for col_idx, col_name in enumerate(headers):
            # Preserve first column value
            if col_idx == 0:
                val = current_row.get(col_name, '')
                if pd.isna(val) or not str(val).strip():
                    val = ''
            else:
                val = self._get_column_value_optimized(col_name, srow, column_mapping)
            
            new_vals.append(self._ensure_scalar(val))
        
        return new_vals

    def _get_column_value_optimized(self, col_name: str, srow: pd.Series, column_mapping: dict) -> str:
        """OPTIMIZED: Get column value with efficient mapping and transformations"""
        
        # Direct column mapping
        if col_name in column_mapping.values():
            src_col = next((src for src, tgt in column_mapping.items() if tgt == col_name), None)
            if src_col and src_col in srow.index:
                cand = srow[src_col]
                if pd.notna(cand) and str(cand).strip() not in ['', '- None -']:
                    # FIXED: Handle date formatting
                    if src_col == 'Start date':
                        parsed_date = self._parse_date_flexible(str(cand))
                        return parsed_date.strftime('%m/%d/%Y') if parsed_date else str(cand).strip()
                    
                    # FIXED: Correct Rent free calculation
                    elif src_col == 'Total months fitout & rent free (for model)' and col_name == 'Rent free':
                        # Calculate rent free = total - fitout
                        try:
                            total_months = float(str(cand).replace(',', '')) if str(cand).replace(',', '').replace('.', '').isdigit() else 0
                            months_fitout = srow.get('Months fit-out (for model)', 0)
                            months_fitout = float(str(months_fitout).replace(',', '')) if str(months_fitout).replace(',', '').replace('.', '').isdigit() else 0
                            rent_free = max(0, total_months - months_fitout)
                            return str(int(rent_free)) if rent_free.is_integer() else str(rent_free)
                        except (ValueError, TypeError):
                            return '0'
                    
                    # Apply column-specific transformations
                    else:
                        return self._transform_column_value(src_col, cand, srow)
        
        # Special field handling
        field_mappings = {
            'Item2': 'Leasing period',
            'Note': 'Committed',
            'Factory code': srow.get('Unit name', ''),
            'Tenant code': srow.get('Tenant ID', ''),
            'Tenant name': srow.get('Tenant', ''),
        }
        
        if col_name in field_mappings:
            return str(field_mappings[col_name]) if pd.notna(field_mappings[col_name]) else ''
        
        # Project code from Phase column
        if col_name == 'Project code':
            phase_info = self._parse_phase_column(str(srow.get('Phase', '')))
            return phase_info.get('project_code', '') or ''
        
        return ''

    def _batch_write_to_excel(self, sheet: xw.Sheet, batch_data: List, headers: List[str]) -> int:
        """OPTIMIZED: Batch write operations to Excel with error handling"""
        if not batch_data:
            return 0
        
        rows_added = 0
        
        try:
            print(f"   ⚡ Writing {len(batch_data)} rows to Excel...")
            
            # OPTIMIZATION: Batch write multiple rows at once when possible
            batch_size = min(20, len(batch_data))  # Process in smaller batches for stability
            
            for i in range(0, len(batch_data), batch_size):
                batch_chunk = batch_data[i:i + batch_size]
                
                for excel_row, vals in batch_chunk:
                    try:
                        # Write entire row at once
                        sheet.range((excel_row, 1), (excel_row, len(headers))).value = vals
                        
                        # Apply formatting if enabled
                        if self._is_feature_enabled('enable_row_highlighting'):
                            self._apply_row_formatting(sheet, excel_row, len(headers), 'add')
                        
                        rows_added += 1
                        
                    except Exception as e:
                        print(f"   ⚠️ Failed to write row {excel_row}: {e}")
                        continue
                
                # Brief pause between batches to prevent COM issues
                if i + batch_size < len(batch_data):
                    import time
                    time.sleep(0.01)
            
            print(f"   ✅ Successfully added {rows_added} new rows")
            
        except Exception as e:
            print(f"   ⚠️ Batch write operation failed: {e}")
        
        return rows_added



    @staticmethod
    def _ensure_scalar(val):
        if hasattr(val, 'iloc') and len(val) > 0:
            return val.iloc[0]
        if hasattr(val, 'item'):
            return val.item()
        if pd.isna(val):
            return ''
        return val

    def _transform_column_value(self, src_col: str, value, srow: pd.Series) -> str:
        """
        FIXED: Transform column values based on specific business logic
        """
        # General empty check for other columns
        if pd.isna(value) or str(value).strip() in ['', '- None -']:
            return ''
        
        if src_col == 'UFL Status':
            ufl_status = str(value).strip().lower()
            return 'Y' if 'handed over' in ufl_status else 'N'
        
        # Payment term logic: Payment term (for model) -> Payment term
        elif src_col == 'Payment term (for model)':
            payment_term = str(value).strip()
            payment_mapping = {
                'Quarterly': '3',
                'Monthly': '1', 
                'Semi-Annual': '6',
                'Yearly': '12'
            }
            
            if payment_term in payment_mapping:
                return payment_mapping[payment_term]
            elif payment_term == 'One-Time Payment':
                # Use Lease period value
                lease_period = srow.get('Lease period', '')
                return str(lease_period) if pd.notna(lease_period) and str(lease_period).strip() else ''
            else:
                return str(value)
        
        # Fitting out logic: Months fit-out (for model) -> fitting out (direct copy)
        elif src_col == 'Months fit-out (for model)':
            return str(value)
        
        # For other columns, return as-is
        else:
            return str(value)
    def _normalize_key(self, value) -> str:
        """Normalize a value for consistent key matching - no caching needed"""
        if pd.isna(value):
            return ""
        return str(value).strip().lower()

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

    def compare_and_highlight_summary_files(self, summary_old_path: str, summary_new_path: str, generate_log: bool = True, log_file_path: str = None) -> bool:
        """
        Compare summary_old (previous period) vs summary_new (current period) Summary files
        and highlight changes in summary_new
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            generate_log: Whether to generate detailed change log file
            log_file_path: Optional custom path for log file
        
        Returns:
            True if comparison and highlighting completed successfully
        """
        try:
            print("🔍 Starting Summary file comparison...")
            
            # Create comparator instance
            comparator = SummaryComparator()
            
            # Perform comparison and highlighting with optional logging
            success = comparator.process_summary_comparison(summary_old_path, summary_new_path, generate_log, log_file_path)
            
            if success:
                print("   ✅ Summary file comparison and highlighting completed!")
            else:
                print("   ❌ Summary file comparison failed")
            
            return success
            
        except Exception as e:
            print(f"   ❌ Error in summary comparison: {e}")
            return False
    
    def apply_summary_highlighting(self, summary_path: str, changed_rows: set) -> bool:
        """
        Apply highlighting to specific rows in Summary file
        
        Args:
            summary_path: Path to Summary file to highlight
            changed_rows: Set of row numbers (1-indexed) to highlight
        
        Returns:
            True if highlighting was successful
        """
        try:
            if not changed_rows:
                print("   ℹ️ No rows to highlight in Summary file")
                return True
                
            print(f"🎨 Applying highlighting to {len(changed_rows)} rows in Summary file")
            
            # Create comparator instance for highlighting functionality
            comparator = SummaryComparator()
            
            # Apply highlighting
            success = comparator.apply_highlighting_to_summary(summary_path, changed_rows)
            
            return success
            
        except Exception as e:
            print(f"   ❌ Error applying summary highlighting: {e}")
            return False
