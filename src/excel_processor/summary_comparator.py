# excel_processor/summary_comparator.py
import pandas as pd
import xlwings as xw
from typing import Dict, List, Tuple, Set
from datetime import datetime
import os

class SummaryComparator:
    """
    Handles comparison between T7 (previous month) and T8 (current month) Summary files
    to identify and highlight changes in the Summary file.
    """
    
    def __init__(self):
        # Key columns for entity matching (predefined entity mapping keys)
        self.entity_key_columns = ['Unit name', 'Tenant ID', 'Tenant']
        
        # Document Number column for primary comparison
        self.document_number_column = 'Document Number'
        
        # Tracked columns for change detection (when comparing same Document Number)
        self.tracked_columns = {
            'GLA': 'GLA',
            'Start date (for model)': 'Start date (for model)', 
            'End date (for model)': 'End date (for model)',
            'Rent USD_Item (for model)': 'Rent USD_Item (for model)',
            'Rent VND_Item (for model)': 'Rent VND_Item (for model)',
            'Escalation rate (for model)': 'Escalation rate (for model)',
            'Service charge (for model)': 'Service charge (for model)',
            'Broker? (Yes/No)': 'Broker? (Yes/No)'
        }

    
    def generate_detailed_change_log(self, summary_old_path: str, summary_new_path: str, log_file_path: str = None) -> str:
        """
        Generate a detailed log file showing row-level changes with specific column modifications
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            log_file_path: Optional custom path for log file. If None, auto-generates based on input files
        
        Returns:
            Path to the generated log file
        """
        import json
        from datetime import datetime
        
        # Auto-generate log file path if not provided
        if log_file_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            old_name = os.path.splitext(os.path.basename(summary_old_path))[0]
            new_name = os.path.splitext(os.path.basename(summary_new_path))[0]
            log_file_path = f"summary_comparison_{old_name}_vs_{new_name}_{timestamp}.log"
        
        print(f"📝 Generating detailed change log: {log_file_path}")
        
        try:
            # Load both files
            old_df = self.load_summary_file(summary_old_path)
            new_df = self.load_summary_file(summary_new_path)
            
            # Create lookup for old data
            old_lookup = {}
            for idx, row in old_df.iterrows():
                key = self.normalize_row_key(row)
                if key and key != '||':
                    old_lookup[key] = row
            
            # Detailed change tracking
            detailed_changes = []
            summary_stats = {
                'total_new_rows': len(new_df),
                'total_old_rows': len(old_df),
                'changed_rows': 0,
                'new_rows': 0,
                'column_change_counts': {}
            }
            
            # Compare each row in new file
            for idx, new_row in new_df.iterrows():
                new_key = self.normalize_row_key(new_row)
                excel_row = idx + 2  # Excel row number (1-indexed + header)
                
                if not new_key or new_key == '||':
                    continue
                
                row_changes = {
                    'excel_row': excel_row,
                    'key_columns': {
                        'Unit name': str(new_row.get('Unit name', '')),
                        'Tenant ID': str(new_row.get('Tenant ID', '')),
                        'Tenant': str(new_row.get('Tenant', ''))
                    },
                    'change_type': '',
                    'modified_columns': []
                }
                
                if new_key in old_lookup:
                    old_row = old_lookup[new_key]
                    has_changes = False
                    
                    # Check each tracked column for changes
                    for new_col, old_col in self.tracked_columns.items():
                        if new_col in new_df.columns and old_col in old_df.columns:
                            new_val = self._normalize_value(new_row.get(new_col), new_col)
                            old_val = self._normalize_value(old_row.get(old_col), old_col)
                            
                            if new_val != old_val:
                                has_changes = True
                                
                                # Track raw values for better readability
                                raw_new_val = str(new_row.get(new_col, ''))
                                raw_old_val = str(old_row.get(old_col, ''))
                                
                                column_change = {
                                    'column': new_col,
                                    'old_value': raw_old_val,
                                    'new_value': raw_new_val,
                                    'normalized_old': old_val,
                                    'normalized_new': new_val
                                }
                                row_changes['modified_columns'].append(column_change)
                                
                                # Update stats
                                if new_col not in summary_stats['column_change_counts']:
                                    summary_stats['column_change_counts'][new_col] = 0
                                summary_stats['column_change_counts'][new_col] += 1
                    
                    if has_changes:
                        row_changes['change_type'] = 'MODIFIED'
                        detailed_changes.append(row_changes)
                        summary_stats['changed_rows'] += 1
                else:
                    # New row
                    row_changes['change_type'] = 'NEW'
                    # Add all column values for new rows
                    for col in self.tracked_columns.keys():
                        if col in new_df.columns:
                            raw_val = str(new_row.get(col, ''))
                            if raw_val.strip():  # Only include non-empty values
                                column_info = {
                                    'column': col,
                                    'old_value': '',
                                    'new_value': raw_val,
                                    'normalized_old': '',
                                    'normalized_new': self._normalize_value(new_row.get(col), col)
                                }
                                row_changes['modified_columns'].append(column_info)
                    
                    detailed_changes.append(row_changes)
                    summary_stats['new_rows'] += 1
            
            # Generate log content
            log_content = self._generate_log_content(
                summary_old_path, summary_new_path, detailed_changes, summary_stats
            )
            
            # Write log file
            with open(log_file_path, 'w', encoding='utf-8') as f:
                f.write(log_content)
            
            print(f"   ✅ Change log generated: {log_file_path}")
            print(f"   📊 Summary: {summary_stats['changed_rows']} modified, {summary_stats['new_rows']} new rows")
            
            return log_file_path
            
        except Exception as e:
            print(f"   ❌ Error generating change log: {e}")
            raise
    
    def _generate_log_content(self, old_path: str, new_path: str, changes: list, stats: dict) -> str:
        """Generate formatted log content"""
        from datetime import datetime
        
        content = []
        content.append("=" * 100)
        content.append("SUMMARY FILE COMPARISON - DETAILED CHANGE LOG")
        content.append("=" * 100)
        content.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        content.append(f"Previous File: {old_path}")
        content.append(f"Current File:  {new_path}")
        content.append("")
        
        # Summary statistics
        content.append("📊 SUMMARY STATISTICS")
        content.append("-" * 50)
        content.append(f"Total rows in previous file: {stats['total_old_rows']}")
        content.append(f"Total rows in current file:  {stats['total_new_rows']}")
        content.append(f"Modified rows: {stats['changed_rows']}")
        content.append(f"New rows:      {stats['new_rows']}")
        content.append(f"Total changes: {stats['changed_rows'] + stats['new_rows']}")
        content.append("")
        
        # Column change statistics
        if stats['column_change_counts']:
            content.append("📈 COLUMN CHANGE FREQUENCY")
            content.append("-" * 50)
            for col, count in sorted(stats['column_change_counts'].items(), key=lambda x: x[1], reverse=True):
                content.append(f"{col}: {count} changes")
            content.append("")
        
        # Detailed changes
        content.append("📝 DETAILED CHANGES")
        content.append("-" * 50)
        
        if not changes:
            content.append("No changes detected.")
        else:
            for i, change in enumerate(changes, 1):
                content.append(f"\n[{i}] ROW {change['excel_row']} - {change['change_type']}")
                content.append(f"    Key: Unit='{change['key_columns']['Unit name']}', "
                             f"TenantID='{change['key_columns']['Tenant ID']}', "
                             f"Tenant='{change['key_columns']['Tenant']}'")
                
                if change['modified_columns']:
                    content.append("    Modified Columns:")
                    for col_change in change['modified_columns']:
                        if change['change_type'] == 'NEW':
                            content.append(f"      • {col_change['column']}: '{col_change['new_value']}'")
                        else:
                            content.append(f"      • {col_change['column']}: '{col_change['old_value']}' → '{col_change['new_value']}'")
                content.append("")
        
        content.append("=" * 100)
        content.append("END OF CHANGE LOG")
        content.append("=" * 100)
        
        return "\n".join(content)
    
    def load_summary_file(self, file_path: str) -> pd.DataFrame:
        """Load summary file with error handling"""
        try:
            # Try different engines based on file extension
            if file_path.endswith('.xlsx'):
                engines = ['openpyxl', 'xlrd']
            elif file_path.endswith('.xls'):
                engines = ['xlrd', 'openpyxl']
            else:
                engines = ['openpyxl', 'xlrd']
            
            for engine in engines:
                try:
                    df = pd.read_excel(file_path, engine=engine)
                    print(f"   Loaded {file_path} using {engine} engine")
                    return df
                except Exception as e:
                    print(f"   Failed to load with {engine}: {e}")
                    continue
            
            raise Exception(f"Failed to load {file_path} with any engine")
            
        except Exception as e:
            print(f"   Error loading {file_path}: {e}")
            raise
    
    def normalize_row_key(self, row: pd.Series) -> str:
        """Create normalized key for row matching - for backward compatibility"""
        key_parts = []
        for col in self.entity_key_columns:
            if col in row:
                value = str(row[col]).strip().lower() if pd.notna(row[col]) else ''
                key_parts.append(value)
            else:
                key_parts.append('')
        return '|'.join(key_parts)

    
    def normalize_entity_key(self, row: pd.Series) -> str:
        """Create normalized key for entity matching using predefined keys"""
        key_parts = []
        for col in self.entity_key_columns:
            if col in row:
                value = str(row[col]).strip().lower() if pd.notna(row[col]) else ''
                key_parts.append(value)
            else:
                key_parts.append('')
        return '|'.join(key_parts)
    
    def get_document_number(self, row: pd.Series) -> str:
        """Extract Document Number from row"""
        if self.document_number_column in row:
            doc_num = str(row[self.document_number_column]).strip() if pd.notna(row[self.document_number_column]) else ''
            return doc_num
        return ''
    
    def compare_summary_files_by_document_number(self, summary_old_path: str, summary_new_path: str) -> Dict[str, any]:
        """
        Compare summary files using Document Number as the primary key identifier.
        
        NEW LOGIC:
        - Use Document Number for comparison (each Document Number is treated as unique record)
        - Highlight newly added Document Numbers (entire rows)
        - Map entities using predefined keys (Unit name, Tenant ID, Tenant)
        - Within each mapped entity, compare the list of Document Numbers
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
        
        Returns:
            Dict with:
            - 'new_document_numbers': set of row indices with newly added Document Numbers
            - 'changed_cells': dict mapping row indices to sets of changed column names
            - 'changed_rows': set of row indices (for backward compatibility)
            - 'entity_mapping': dict mapping entities between old and new files
        """
        print(f"🔧 Comparing Summary files using Document Number logic:")
        print(f"   Previous: {summary_old_path}")
        print(f"   Current:  {summary_new_path}")
        
        # Load both files
        old_df = self.load_summary_file(summary_old_path)
        new_df = self.load_summary_file(summary_new_path)
        
        print(f"   Previous rows: {len(old_df)}, Current rows: {len(new_df)}")
        
        # Check if Document Number column exists
        if self.document_number_column not in new_df.columns:
            print(f"   ⚠️ Warning: '{self.document_number_column}' column not found in new file. Using fallback logic.")
            return self.compare_summary_files(summary_old_path, summary_new_path)
        
        if self.document_number_column not in old_df.columns:
            print(f"   ⚠️ Warning: '{self.document_number_column}' column not found in old file. Treating all as new.")
            # All rows in new file are considered new
            new_document_numbers = set(range(2, len(new_df) + 2))  # Excel row numbers
            return {
                'new_document_numbers': new_document_numbers,
                'changed_cells': {},
                'changed_rows': new_document_numbers,
                'entity_mapping': {}
            }
        
        # Create lookups
        # 1. Document Number lookup for old file
        old_doc_lookup = {}
        for idx, row in old_df.iterrows():
            doc_num = self.get_document_number(row)
            if doc_num:
                old_doc_lookup[doc_num] = row
        
        # 2. Entity mapping lookup for old file
        old_entity_lookup = {}
        for idx, row in old_df.iterrows():
            entity_key = self.normalize_entity_key(row)
            if entity_key and entity_key != '||':
                if entity_key not in old_entity_lookup:
                    old_entity_lookup[entity_key] = []
                doc_num = self.get_document_number(row)
                if doc_num:
                    old_entity_lookup[entity_key].append(doc_num)
        
        print(f"   Previous Document Numbers: {len(old_doc_lookup)}")
        print(f"   Previous Entities: {len(old_entity_lookup)}")
        
        # Track results
        new_document_numbers = set()
        changed_cells = {}
        changed_rows = set()
        entity_mapping = {}
        
        # Compare each row in new file
        for idx, new_row in new_df.iterrows():
            row_num = idx + 2  # Excel row number (1-indexed + header)
            doc_num = self.get_document_number(new_row)
            entity_key = self.normalize_entity_key(new_row)
            
            if not doc_num:
                continue
            
            # Check if this Document Number exists in old file
            if doc_num in old_doc_lookup:
                # Document Number exists - check for changes in tracked columns
                old_row = old_doc_lookup[doc_num]
                changed_columns = set()
                
                for new_col, old_col in self.tracked_columns.items():
                    if new_col in new_df.columns and old_col in old_df.columns:
                        new_val = self._normalize_value(new_row.get(new_col), new_col)
                        old_val = self._normalize_value(old_row.get(old_col), old_col)
                        
                        if new_val != old_val:
                            changed_columns.add(new_col)
                
                if changed_columns:
                    changed_rows.add(row_num)
                    changed_cells[row_num] = changed_columns
                    print(f"   📝 Document Number '{doc_num}' has changes in row {row_num}")
            
            else:
                # New Document Number - highlight entire row
                new_document_numbers.add(row_num)
                changed_rows.add(row_num)
                
                # Mark all tracked columns as new for highlighting
                all_columns = set(self.tracked_columns.keys())
                # Only include columns that actually exist in the new dataframe
                all_columns = {col for col in all_columns if col in new_df.columns}
                changed_cells[row_num] = all_columns
                
                print(f"   🆕 New Document Number '{doc_num}' in row {row_num}")
                
                # Track entity mapping for this new Document Number
                if entity_key and entity_key != '||':
                    if entity_key not in entity_mapping:
                        entity_mapping[entity_key] = {
                            'entity_info': {
                                'Unit name': str(new_row.get('Unit name', '')),
                                'Tenant ID': str(new_row.get('Tenant ID', '')),
                                'Tenant': str(new_row.get('Tenant', ''))
                            },
                            'old_document_numbers': old_entity_lookup.get(entity_key, []),
                            'new_document_numbers': []
                        }
                    entity_mapping[entity_key]['new_document_numbers'].append(doc_num)
        
        print(f"   🎯 Found {len(new_document_numbers)} newly added Document Numbers")
        print(f"   📊 Found {len(changed_rows)} total rows with changes")
        print(f"   🏢 Mapped {len(entity_mapping)} entities with new Document Numbers")
        
        return {
            'new_document_numbers': new_document_numbers,
            'changed_cells': changed_cells,
            'changed_rows': changed_rows,
            'entity_mapping': entity_mapping
        }
    
    def compare_summary_files(self, summary_old_path: str, summary_new_path: str) -> Dict[str, any]:
        """
        Compare summary_old vs summary_new and return specific cells that changed
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
        
        Returns:
            Dict with 'changed_cells': dict mapping row indices to sets of changed column names,
            and 'changed_rows': set of row indices (for backward compatibility)
        """
        print(f"Comparing Summary files:")
        print(f"   Previous: {summary_old_path}")
        print(f"   Current:  {summary_new_path}")
        
        # Load both files
        old_df = self.load_summary_file(summary_old_path)
        new_df = self.load_summary_file(summary_new_path)
        
        print(f"   Previous rows: {len(old_df)}, Current rows: {len(new_df)}")
        
        # Create lookup for old data
        old_lookup = {}
        for idx, row in old_df.iterrows():
            key = self.normalize_row_key(row)
            if key and key != '||':  # Skip empty keys
                old_lookup[key] = row
        
        print(f"   Previous lookup created with {len(old_lookup)} unique keys")
        
        changed_rows = set()
        changed_cells = {}  # row_num -> set of changed column names
        
        # Compare each row in new file with corresponding row in old file
        for idx, new_row in new_df.iterrows():
            new_key = self.normalize_row_key(new_row)
            row_num = idx + 2  # +2 for Excel row number (1-indexed + header)
            
            if not new_key or new_key == '||':
                continue
                
            if new_key in old_lookup:
                old_row = old_lookup[new_key]
                
                # Check each tracked column for changes
                changed_columns = set()
                for new_col, old_col in self.tracked_columns.items():
                    if new_col in new_df.columns and old_col in old_df.columns:
                        new_val = self._normalize_value(new_row.get(new_col), new_col)
                        old_val = self._normalize_value(old_row.get(old_col), old_col)
                        
                        if new_val != old_val:
                            changed_columns.add(new_col)
                
                if changed_columns:
                    changed_rows.add(row_num)
                    changed_cells[row_num] = changed_columns
            else:
                # New row in current file - mark all tracked columns as changed
                changed_rows.add(row_num)
                changed_columns = set(self.tracked_columns.keys())
                # Only include columns that actually exist in the new dataframe
                changed_columns = {col for col in changed_columns if col in new_df.columns}
                changed_cells[row_num] = changed_columns
                print(f"   Row {row_num} is new in current file")
        
        print(f"   Found {len(changed_rows)} changed/new rows in current file")
        print(f"   Found {sum(len(cols) for cols in changed_cells.values())} individual cell changes")
        
        return {
            'changed_rows': changed_rows,
            'changed_cells': changed_cells
        }

    def compare_summary_files_with_oldest_match(self, summary_old_path: str, summary_new_path: str) -> Dict[str, any]:
        """
        Compare each row in new Excel file against oldest matching record in old file
        using key mapping ['Unit name', 'Tenant ID', 'Tenant'] and Date created column.
        
        For Document Numbers that don't exist in old file, find oldest record with same key mapping.
        Highlight only changed cells (light blue) or entire new rows (yellow).
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
        
        Returns:
            Dict with 'changed_cells': dict mapping row indices to sets of changed column names,
            'changed_rows': set of row indices, and 'new_rows': set of entirely new row indices
        """
        print(f"Comparing Summary files with oldest match logic:")
        print(f"   Previous: {summary_old_path}")
        print(f"   Current:  {summary_new_path}")
        
        # Load both files
        old_df = self.load_summary_file(summary_old_path)
        new_df = self.load_summary_file(summary_new_path)
        
        print(f"   Previous rows: {len(old_df)}, Current rows: {len(new_df)}")
        
        # Ensure Date created column exists
        date_created_col = 'Date created'
        if date_created_col not in old_df.columns:
            print(f"   Warning: '{date_created_col}' column not found in old file")
            # Fallback to regular comparison
            return self.compare_summary_files(summary_old_path, summary_new_path)
        
        # Group old records by entity key and find oldest in each group
        old_grouped = {}
        for idx, row in old_df.iterrows():
            entity_key = self.normalize_entity_key(row)
            if not entity_key or entity_key == '||':
                continue
                
            # Parse date created
            date_created = self._parse_date_created(row.get(date_created_col))
            if date_created is None:
                continue
                
            # Store all records for this entity key
            if entity_key not in old_grouped:
                old_grouped[entity_key] = []
            old_grouped[entity_key].append({
                'row': row,
                'date_created': date_created,
                'document_number': self.get_document_number(row)
            })
        
        # Find oldest record for each entity key
        old_oldest_lookup = {}
        old_document_lookup = {}  # For direct document number matches
        
        for entity_key, records in old_grouped.items():
            # Sort by date created to find oldest
            records.sort(key=lambda x: x['date_created'])
            oldest_record = records[0]
            old_oldest_lookup[entity_key] = oldest_record['row']
            
            # Also create document number lookup for direct matches
            for record in records:
                doc_num = record['document_number']
                if doc_num:
                    old_document_lookup[doc_num] = record['row']
        
        print(f"   Created lookup with {len(old_oldest_lookup)} entity groups")
        print(f"   Created document lookup with {len(old_document_lookup)} documents")
        
        changed_rows = set()
        changed_cells = {}  # row_num -> set of changed column names
        new_rows = set()    # Entirely new rows (no matching entity key)
        
        # Compare each row in new file
        for idx, new_row in new_df.iterrows():
            row_num = idx + 2  # +2 for Excel row number (1-indexed + header)
            new_doc_num = self.get_document_number(new_row)
            entity_key = self.normalize_entity_key(new_row)
            
            if not entity_key or entity_key == '||':
                continue
            
            # First, try to find direct document number match
            comparison_row = None
            if new_doc_num and new_doc_num in old_document_lookup:
                comparison_row = old_document_lookup[new_doc_num]
                print(f"   Row {row_num}: Found direct document match for '{new_doc_num}'")
            
            # If no direct document match, use oldest record with same entity key
            elif entity_key in old_oldest_lookup:
                comparison_row = old_oldest_lookup[entity_key]
                print(f"   Row {row_num}: Using oldest record for entity key (doc '{new_doc_num}' not found)")
            
            # If no matching entity key at all, mark as entirely new
            else:
                new_rows.add(row_num)
                print(f"   Row {row_num}: Entirely new entity key")
                continue
            
            # Compare cells between new row and comparison row
            changed_columns = set()
            for new_col, old_col in self.tracked_columns.items():
                if new_col in new_df.columns and old_col in old_df.columns:
                    new_val = self._normalize_value(new_row.get(new_col), new_col)
                    old_val = self._normalize_value(comparison_row.get(old_col), old_col)
                    
                    if new_val != old_val:
                        changed_columns.add(new_col)
            
            if changed_columns:
                changed_rows.add(row_num)
                changed_cells[row_num] = changed_columns
        
        print(f"   Found {len(changed_rows)} rows with cell changes")
        print(f"   Found {len(new_rows)} entirely new rows")
        print(f"   Found {sum(len(cols) for cols in changed_cells.values())} individual cell changes")
        
        return {
            'changed_rows': changed_rows,
            'changed_cells': changed_cells,
            'new_rows': new_rows
        }
    
    def _parse_date_created(self, date_value) -> datetime:
        """Parse date created value to datetime object for comparison"""
        if pd.isna(date_value):
            return None
            
        if isinstance(date_value, datetime):
            return date_value
            
        try:
            date_str = str(date_value).strip()
            
            # Try common date formats
            for date_format in ['%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d']:
                try:
                    return datetime.strptime(date_str, date_format)
                except ValueError:
                    continue
            
            # If no format works, return None
            return None
            
        except Exception:
            return None
    
    def _normalize_value(self, value, column_name: str = '') -> str:
        """Normalize value for comparison based on column type"""
        if pd.isna(value):
            return ''
        
        # Handle datetime objects
        if isinstance(value, datetime):
            return value.strftime('%m/%d/%Y')
        
        # Convert to string and normalize
        str_val = str(value).strip()
        
        # Handle date columns
        if 'date' in column_name.lower():
            try:
                # Try to parse and reformat date
                if '/' in str_val:
                    # Try mm/dd/yyyy format
                    dt = datetime.strptime(str_val, '%m/%d/%Y')
                elif '-' in str_val:
                    # Try yyyy-mm-dd format
                    dt = datetime.strptime(str_val, '%Y-%m-%d')
                else:
                    return str_val.lower()
                return dt.strftime('%m/%d/%Y')
            except:
                return str_val.lower()
        
        # Handle numeric columns (Rent, Service charge, Escalation rate)
        elif any(keyword in column_name.lower() for keyword in ['rent', 'service charge', 'escalation rate']):
            try:
                # Remove common formatting characters
                clean_val = str_val.replace(',', '').replace('$', '').replace('%', '').strip()
                if clean_val and clean_val.replace('.', '').replace('-', '').isdigit():
                    # Convert to float for consistent comparison
                    num_val = float(clean_val)
                    return str(num_val)
                else:
                    return str_val.lower()
            except:
                return str_val.lower()
        
        # Handle boolean columns (Broker)
        elif 'broker' in column_name.lower():
            normalized = str_val.lower()
            # Normalize common boolean representations
            if normalized in ['yes', 'y', 'true', '1']:
                return 'yes'
            elif normalized in ['no', 'n', 'false', '0']:
                return 'no'
            else:
                return normalized
        
        # Handle GLA (numeric)
        elif 'gla' in column_name.lower():
            try:
                clean_val = str_val.replace(',', '').strip()
                if clean_val and clean_val.replace('.', '').isdigit():
                    return str(float(clean_val))
                else:
                    return str_val.lower()
            except:
                return str_val.lower()
        
        # Default: return normalized string
        return str_val.lower()
    
    def apply_highlighting_to_summary(self, summary_path: str, comparison_result: Dict[str, any]) -> bool:
        """
        Apply highlighting to specific changed cells in the Summary file
        
        Args:
            summary_path: Path to the T8 Summary file
            comparison_result: Result from compare_summary_files containing changed_cells and changed_rows
        
        Returns:
            True if highlighting was successful, False otherwise
        """
        changed_cells = comparison_result.get('changed_cells', {})
        changed_rows = comparison_result.get('changed_rows', set())
        
        if not changed_cells and not changed_rows:
            print("   No cells to highlight in Summary file")
            return True
        
        try:
            print(f"Applying cell-specific highlighting to Summary file")
            print(f"   Rows with changes: {len(changed_rows)}")
            print(f"   Total changed cells: {sum(len(cols) for cols in changed_cells.values())}")
            
            # Open the file with xlwings
            app = xw.App(visible=False, add_book=False)
            try:
                wb = app.books.open(summary_path)
                sheet = wb.sheets[0]  # Assume first sheet
                
                # Get column mapping for the sheet
                header_row = sheet.range('1:1').value
                if not header_row:
                    print("   Could not read header row")
                    return False
                
                # Create column name to column number mapping
                col_mapping = {}
                for col_idx, col_name in enumerate(header_row):
                    if col_name:
                        col_mapping[str(col_name).strip()] = col_idx + 1
                
                highlighted_cells = 0
                for row_num, changed_column_names in changed_cells.items():
                    try:
                        for col_name in changed_column_names:
                            # Find the column number for this column name
                            col_num = None
                            
                            # Direct match first
                            if col_name in col_mapping:
                                col_num = col_mapping[col_name]
                            else:
                                # Try fuzzy matching for column names
                                for header_col, header_col_num in col_mapping.items():
                                    if col_name.lower().strip() in header_col.lower().strip() or \
                                       header_col.lower().strip() in col_name.lower().strip():
                                        col_num = header_col_num
                                        break
                            
                            if col_num:
                                # Highlight the specific cell with light blue background
                                cell = sheet.range((row_num, col_num))
                                cell.color = (173, 216, 230)  # Light blue for changed cells
                                highlighted_cells += 1
                            else:
                                print(f"   Could not find column '{col_name}' in header row")
                    
                    except Exception as e:
                        print(f"   Failed to highlight cells in row {row_num}: {e}")
                
                # Save the file
                wb.save()
                print(f"   Saved Summary file with {highlighted_cells} highlighted cells")
                
                return True
                
            finally:
                wb.close()
                app.quit()
                
        except Exception as e:
            print(f"   Error applying highlighting to Summary file: {e}")
            return False

    def apply_enhanced_highlighting_to_summary(self, summary_path: str, comparison_result: Dict[str, any]) -> bool:
        """
        Apply enhanced highlighting to Summary file with different colors for different change types:
        - Light blue for changed cells
        - Yellow for entirely new rows (no matching entity key)
        
        Args:
            summary_path: Path to the T8 Summary file
            comparison_result: Result from compare_summary_files_with_oldest_match
        
        Returns:
            True if highlighting was successful, False otherwise
        """
        changed_cells = comparison_result.get('changed_cells', {})
        changed_rows = comparison_result.get('changed_rows', set())
        new_rows = comparison_result.get('new_rows', set())
        
        if not changed_cells and not changed_rows and not new_rows:
            print("   No cells to highlight in Summary file")
            return True
        
        try:
            print(f"Applying enhanced highlighting to Summary file")
            print(f"   Rows with cell changes: {len(changed_rows)}")
            print(f"   Entirely new rows: {len(new_rows)}")
            print(f"   Total changed cells: {sum(len(cols) for cols in changed_cells.values())}")
            
            # Open the file with xlwings
            app = xw.App(visible=False, add_book=False)
            try:
                wb = app.books.open(summary_path)
                sheet = wb.sheets[0]  # Assume first sheet
                
                # Get column mapping for the sheet
                header_row = sheet.range('1:1').value
                if not header_row:
                    print("   Could not read header row")
                    return False
                
                # Create column name to column number mapping
                col_mapping = {}
                for col_idx, col_name in enumerate(header_row):
                    if col_name:
                        col_mapping[str(col_name).strip()] = col_idx + 1
                
                highlighted_cells = 0
                highlighted_rows = 0
                
                # Highlight changed cells with light blue
                for row_num, changed_column_names in changed_cells.items():
                    try:
                        for col_name in changed_column_names:
                            # Find the column number for this column name
                            col_num = None
                            
                            # Direct match first
                            if col_name in col_mapping:
                                col_num = col_mapping[col_name]
                            else:
                                # Try fuzzy matching for column names
                                for header_col, header_col_num in col_mapping.items():
                                    if col_name.lower().strip() in header_col.lower().strip() or \
                                       header_col.lower().strip() in col_name.lower().strip():
                                        col_num = header_col_num
                                        break
                            
                            if col_num:
                                # Highlight the specific cell with light blue background
                                cell = sheet.range((row_num, col_num))
                                cell.color = (173, 216, 230)  # Light blue for changed cells
                                highlighted_cells += 1
                            else:
                                print(f"   Could not find column '{col_name}' in header row")
                    
                    except Exception as e:
                        print(f"   Failed to highlight cells in row {row_num}: {e}")
                
                # Highlight entirely new rows with yellow
                for row_num in new_rows:
                    try:
                        # Highlight the entire row with yellow background
                        last_col = len(header_row)
                        row_range = sheet.range((row_num, 1), (row_num, last_col))
                        row_range.color = (255, 255, 0)  # Yellow for entirely new rows
                        highlighted_rows += 1
                        
                    except Exception as e:
                        print(f"   Failed to highlight row {row_num}: {e}")
                
                # Save the file
                wb.save()
                print(f"   Saved Summary file with {highlighted_cells} highlighted cells and {highlighted_rows} highlighted rows")
                
                return True
                
            finally:
                wb.close()
                app.quit()
                
        except Exception as e:
            print(f"   Error applying enhanced highlighting to Summary file: {e}\"")
            return False

    
    def apply_document_number_highlighting_to_summary(self, summary_path: str, comparison_result: Dict[str, any]) -> bool:
        """
        Apply highlighting to newly added Document Numbers (entire rows) in the Summary file
        
        Args:
            summary_path: Path to the T8 Summary file
            comparison_result: Result from compare_summary_files_by_document_number
        
        Returns:
            True if highlighting was successful, False otherwise
        """
        new_document_numbers = comparison_result.get('new_document_numbers', set())
        
        if not new_document_numbers:
            print("   No new document numbers to highlight in Summary file")
            return True
        
        try:
            print(f"🎨 Applying Document Number-based highlighting to Summary file")
            print(f"   Newly added Document Numbers: {len(new_document_numbers)}")
            
            # Open the file with xlwings
            app = xw.App(visible=False, add_book=False)
            try:
                wb = app.books.open(summary_path)
                sheet = wb.sheets[0]  # Assume first sheet
                
                # Get column mapping for the sheet
                header_row = sheet.range('1:1').value
                if not header_row:
                    print("   Could not read header row")
                    return False
                
                highlighted_rows = 0
                
                # Highlight newly added Document Numbers (entire rows) only
                for row_num in new_document_numbers:
                    try:
                        # Highlight entire row with light yellow background for new Document Numbers
                        row_range = sheet.range((row_num, 1), (row_num, len(header_row)))
                        row_range.color = (255, 255, 180)  # Light yellow for new Document Numbers
                        highlighted_rows += 1
                        print(f"   🆕 Highlighted entire row {row_num} (new Document Number)")
                    except Exception as e:
                        print(f"   Failed to highlight row {row_num}: {e}")
                
                # Save the file
                wb.save()
                print(f"   ✅ Saved Summary file with {highlighted_rows} highlighted rows")
                
                return True
                
            finally:
                wb.close()
                app.quit()
                
        except Exception as e:
            print(f"   ❌ Error applying Document Number highlighting to Summary file: {e}")
            return False

    
    def process_summary_comparison_with_document_numbers(self, summary_old_path: str, summary_new_path: str, generate_log: bool = True, log_file_path: str = None) -> bool:
        """
        Complete workflow using NEW Document Number-based comparison logic
        
        This method implements the new requirements:
        1. Use Document Number as the key identifier
        2. Highlight newly added Document Numbers (entire rows)
        3. Map entities using predefined keys
        4. Simplify review process by only flagging genuinely new entries
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            generate_log: Whether to generate detailed change log file
            log_file_path: Optional custom path for log file
        
        Returns:
            True if process completed successfully
        """
        try:
            # Validate files exist
            if not os.path.exists(summary_old_path):
                print(f"   Previous file not found: {summary_old_path}")
                return False
            
            if not os.path.exists(summary_new_path):
                print(f"   Current file not found: {summary_new_path}")
                return False
            
            print(f"🔧 Starting Document Number-based comparison workflow")
            
            # Use the new Document Number-based comparison logic
            comparison_results = self.compare_summary_files_by_document_number(summary_old_path, summary_new_path)
            
            new_document_numbers = comparison_results['new_document_numbers']
            changed_rows = comparison_results['changed_rows']
            changed_cells = comparison_results['changed_cells']
            entity_mapping = comparison_results['entity_mapping']
            
            # Generate detailed log if requested (using new format)
            if generate_log:
                log_path = self.generate_document_number_change_log(
                    summary_old_path, summary_new_path, comparison_results, log_file_path
                )
                print(f"   📝 Detailed log saved: {log_path}")
            
            # Apply Document Number-based highlighting
            success = self.apply_document_number_highlighting_to_summary(summary_new_path, comparison_results)
            
            if success:
                print(f"   🎉 Document Number-based comparison completed successfully!")
                print(f"   🆕 Newly added Document Numbers: {len(new_document_numbers)}")
                print(f"   📊 Total rows with changes: {len(changed_rows)}")
                print(f"   🏢 Entities with new Document Numbers: {len(entity_mapping)}")
                if generate_log:
                    print(f"   📝 Detailed change log: {log_path}")
            
            return success
            
        except Exception as e:
            print(f"   ❌ Error in Document Number-based comparison process: {e}")
            return False

    
    def generate_document_number_change_log(self, summary_old_path: str, summary_new_path: str, comparison_results: Dict, log_file_path: str = None) -> str:
        """
        Generate a detailed log file for Document Number-based comparison
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            comparison_results: Results from compare_summary_files_by_document_number
            log_file_path: Optional custom path for log file
        
        Returns:
            Path to the generated log file
        """
        from datetime import datetime
        
        # Auto-generate log file path if not provided
        if log_file_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            old_name = os.path.splitext(os.path.basename(summary_old_path))[0]
            new_name = os.path.splitext(os.path.basename(summary_new_path))[0]
            log_file_path = f"document_number_comparison_{old_name}_vs_{new_name}_{timestamp}.log"
        
        print(f"📝 Generating Document Number-based change log: {log_file_path}")
        
        try:
            # Load files for detailed analysis
            old_df = self.load_summary_file(summary_old_path)
            new_df = self.load_summary_file(summary_new_path)
            
            new_document_numbers = comparison_results['new_document_numbers']
            entity_mapping = comparison_results['entity_mapping']
            changed_cells = comparison_results['changed_cells']
            
            # Generate log content
            content = []
            content.append("=" * 100)
            content.append("DOCUMENT NUMBER-BASED SUMMARY COMPARISON LOG")
            content.append("=" * 100)
            content.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            content.append(f"Previous File: {summary_old_path}")
            content.append(f"Current File:  {summary_new_path}")
            content.append("")
            
            # Summary statistics
            content.append("📊 SUMMARY STATISTICS")
            content.append("-" * 50)
            content.append(f"Total rows in previous file: {len(old_df)}")
            content.append(f"Total rows in current file:  {len(new_df)}")
            content.append(f"Newly added Document Numbers: {len(new_document_numbers)}")
            content.append(f"Entities with new Document Numbers: {len(entity_mapping)}")
            content.append(f"Total rows with changes: {len(changed_cells)}")
            content.append("")
            
            # Entity mapping summary
            if entity_mapping:
                content.append("🏢 ENTITY MAPPING SUMMARY")
                content.append("-" * 50)
                for entity_key, mapping_info in entity_mapping.items():
                    entity_info = mapping_info['entity_info']
                    old_docs = mapping_info['old_document_numbers']
                    new_docs = mapping_info['new_document_numbers']
                    
                    content.append(f"Entity: {entity_info['Unit name']} | {entity_info['Tenant ID']} | {entity_info['Tenant']}")
                    content.append(f"  Previous Document Numbers: {len(old_docs)} ({', '.join(old_docs) if old_docs else 'None'})")
                    content.append(f"  New Document Numbers: {len(new_docs)} ({', '.join(new_docs)})")
                    content.append("")
            
            # Detailed new Document Numbers
            content.append("🆕 NEWLY ADDED DOCUMENT NUMBERS")
            content.append("-" * 50)
            
            if not new_document_numbers:
                content.append("No new Document Numbers found.")
            else:
                for i, row_num in enumerate(sorted(new_document_numbers), 1):
                    try:
                        # Get row data (convert from Excel row number to DataFrame index)
                        df_idx = row_num - 2
                        if df_idx < len(new_df):
                            row_data = new_df.iloc[df_idx]
                            doc_num = self.get_document_number(row_data)
                            
                            content.append(f"[{i}] ROW {row_num} - NEW DOCUMENT NUMBER")
                            content.append(f"    Document Number: '{doc_num}'")
                            content.append(f"    Entity: Unit='{row_data.get('Unit name', '')}', "
                                         f"TenantID='{row_data.get('Tenant ID', '')}', "
                                         f"Tenant='{row_data.get('Tenant', '')}'")
                            
                            # Show key values for new Document Number
                            content.append("    Key Values:")
                            for col in self.tracked_columns.keys():
                                if col in new_df.columns:
                                    value = str(row_data.get(col, ''))
                                    if value.strip():
                                        content.append(f"      • {col}: '{value}'")
                            content.append("")
                    except Exception as e:
                        content.append(f"[{i}] ROW {row_num} - Error reading data: {e}")
                        content.append("")
            
            # Changes in existing Document Numbers
            existing_doc_changes = {row: cols for row, cols in changed_cells.items() 
                                  if row not in new_document_numbers}
            
            if existing_doc_changes:
                content.append("📝 CHANGES IN EXISTING DOCUMENT NUMBERS")
                content.append("-" * 50)
                
                for i, (row_num, changed_columns) in enumerate(sorted(existing_doc_changes.items()), 1):
                    try:
                        # Get row data
                        df_idx = row_num - 2
                        if df_idx < len(new_df):
                            row_data = new_df.iloc[df_idx]
                            doc_num = self.get_document_number(row_data)
                            
                            content.append(f"[{i}] ROW {row_num} - MODIFIED DOCUMENT NUMBER")
                            content.append(f"    Document Number: '{doc_num}'")
                            content.append(f"    Entity: Unit='{row_data.get('Unit name', '')}', "
                                         f"TenantID='{row_data.get('Tenant ID', '')}', "
                                         f"Tenant='{row_data.get('Tenant', '')}'")
                            content.append(f"    Modified Columns: {', '.join(changed_columns)}")
                            content.append("")
                    except Exception as e:
                        content.append(f"[{i}] ROW {row_num} - Error reading data: {e}")
                        content.append("")
            
            content.append("=" * 100)
            content.append("END OF DOCUMENT NUMBER COMPARISON LOG")
            content.append("=" * 100)
            
            # Write log file
            with open(log_file_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(content))
            
            print(f"   ✅ Document Number change log generated: {log_file_path}")
            
            return log_file_path
            
        except Exception as e:
            print(f"   ❌ Error generating Document Number change log: {e}")
            raise
    
    def process_summary_comparison(self, summary_old_path: str, summary_new_path: str, generate_log: bool = True, log_file_path: str = None) -> bool:
        """
        Complete workflow: compare summary_old vs summary_new and highlight changes in summary_new
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            generate_log: Whether to generate detailed change log file
            log_file_path: Optional custom path for log file
        
        Returns:
            True if process completed successfully
        """
        try:
            # Validate files exist
            if not os.path.exists(summary_old_path):
                print(f"   Previous file not found: {summary_old_path}")
                return False
            
            if not os.path.exists(summary_new_path):
                print(f"   Current file not found: {summary_new_path}")
                return False
            
            # Generate detailed log if requested
            if generate_log:
                log_path = self.generate_detailed_change_log(summary_old_path, summary_new_path, log_file_path)
                print(f"   Detailed log saved: {log_path}")
            
            # Compare files for highlighting (now returns detailed comparison results)
            comparison_results = self.compare_summary_files(summary_old_path, summary_new_path)
            changed_rows = comparison_results['changed_rows']
            changed_cells = comparison_results['changed_cells']
            
            # Apply cell-specific highlighting
            success = self.apply_highlighting_to_summary(summary_new_path, comparison_results)
            
            if success:
                print(f"   Summary comparison completed successfully!")
                print(f"   Total rows with changes: {len(changed_rows)}")
                print(f"   Total individual cells highlighted: {sum(len(cols) for cols in changed_cells.values())}")
                if generate_log:
                    print(f"   Detailed change log: {log_path}")
            
            return success
            
        except Exception as e:
            print(f"   Error in summary comparison process: {e}")
            return False

    def process_enhanced_summary_comparison(self, summary_old_path: str, summary_new_path: str, 
                                          output_dir: str = None) -> bool:
        """
        Process summary comparison with enhanced logic using oldest matching records
        and cell-level highlighting with different colors for different change types.
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file  
            output_dir: Optional output directory for change logs
        
        Returns:
            True if comparison was successful, False otherwise
        """
        try:
            print(f"\n=== Enhanced Summary Comparison ===")
            
            # Perform enhanced comparison
            comparison_result = self.compare_summary_files_with_oldest_match(
                summary_old_path, summary_new_path
            )
            
            # Apply enhanced highlighting
            highlighting_success = self.apply_enhanced_highlighting_to_summary(
                summary_new_path, comparison_result
            )
            
            if not highlighting_success:
                print("   Warning: Failed to apply highlighting")
            
            # Generate change log if output directory is provided
            if output_dir:
                log_success = self.generate_enhanced_change_log(
                    summary_old_path, summary_new_path, comparison_result, output_dir
                )
                if not log_success:
                    print("   Warning: Failed to generate change log")
            
            return True
            
        except Exception as e:
            print(f"   Error in enhanced summary comparison: {e}")
            return False
    
    def generate_enhanced_change_log(self, summary_old_path: str, summary_new_path: str, 
                                   comparison_result: Dict[str, any], output_dir: str) -> bool:
        """
        Generate detailed change log for enhanced comparison results.
        
        Args:
            summary_old_path: Path to previous period Summary file
            summary_new_path: Path to current period Summary file
            comparison_result: Result from compare_summary_files_with_oldest_match
            output_dir: Directory to save the change log
        
        Returns:
            True if log generation was successful, False otherwise
        """
        try:
            # Load the new file to get row data
            new_df = self.load_summary_file(summary_new_path)
            
            # Create timestamp for log filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_filename = f"enhanced_summary_changes_{timestamp}.log"
            log_path = os.path.join(output_dir, log_filename)
            
            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)
            
            with open(log_path, 'w', encoding='utf-8') as log_file:
                log_file.write(f"Enhanced Summary Comparison Change Log\n")
                log_file.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                log_file.write(f"Previous File: {summary_old_path}\n")
                log_file.write(f"Current File: {summary_new_path}\n")
                log_file.write(f"Comparison Method: Oldest matching record by entity key\n")
                log_file.write(f"Entity Key Columns: {', '.join(self.entity_key_columns)}\n")
                log_file.write("=" * 80 + "\n\n")
                
                changed_cells = comparison_result.get('changed_cells', {})
                new_rows = comparison_result.get('new_rows', set())
                
                # Log cell-level changes
                if changed_cells:
                    log_file.write(f"CELL-LEVEL CHANGES ({len(changed_cells)} rows):\n")
                    log_file.write("-" * 50 + "\n")
                    
                    for row_num in sorted(changed_cells.keys()):
                        if row_num - 2 < len(new_df):  # Convert back to 0-indexed
                            row_data = new_df.iloc[row_num - 2]
                            entity_key = self.normalize_entity_key(row_data)
                            doc_num = self.get_document_number(row_data)
                            
                            log_file.write(f"Row {row_num}: {entity_key}\n")
                            log_file.write(f"  Document Number: {doc_num or 'N/A'}\n")
                            log_file.write(f"  Changed Columns: {', '.join(sorted(changed_cells[row_num]))}\n")
                            log_file.write(f"  Highlighting: Light blue cells\n\n")
                
                # Log entirely new rows
                if new_rows:
                    log_file.write(f"ENTIRELY NEW ROWS ({len(new_rows)} rows):\n")
                    log_file.write("-" * 50 + "\n")
                    
                    for row_num in sorted(new_rows):
                        if row_num - 2 < len(new_df):  # Convert back to 0-indexed
                            row_data = new_df.iloc[row_num - 2]
                            entity_key = self.normalize_entity_key(row_data)
                            doc_num = self.get_document_number(row_data)
                            
                            log_file.write(f"Row {row_num}: {entity_key}\n")
                            log_file.write(f"  Document Number: {doc_num or 'N/A'}\n")
                            log_file.write(f"  Status: No matching entity key in previous file\n")
                            log_file.write(f"  Highlighting: Yellow row\n\n")
                
                # Summary
                total_changes = len(changed_cells) + len(new_rows)
                total_cell_changes = sum(len(cols) for cols in changed_cells.values())
                
                log_file.write("SUMMARY:\n")
                log_file.write("-" * 20 + "\n")
                log_file.write(f"Total affected rows: {total_changes}\n")
                log_file.write(f"Rows with cell changes: {len(changed_cells)}\n")
                log_file.write(f"Entirely new rows: {len(new_rows)}\n")
                log_file.write(f"Individual cell changes: {total_cell_changes}\n")
            
            print(f"   Enhanced change log saved to: {log_path}")
            return True
            
        except Exception as e:
            print(f"   Error generating enhanced change log: {e}")
            return False
