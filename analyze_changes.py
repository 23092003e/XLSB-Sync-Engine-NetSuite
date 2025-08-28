import pandas as pd

def normalize_value(value):
    """Normalize value for comparison"""
    if pd.isna(value):
        return ''
    return str(value).strip().lower()

# Read files
old_df = pd.read_excel('data/summary/old.xlsx')
new_df = pd.read_excel('data/summary/new_28.xlsx')

# Column mapping between new and old files
column_mapping = {
    'GLA': 'GLA',
    'Start date (for model)': 'Start date (for model)',
    'End date (for model)': 'End date (for model)', 
    'Rent (VND)_Item': 'Rent VND_Item (for model)',
    'Rent (USD)_Item': 'Rent USD_Item (for model)', 
    'Escalation rate': 'Escalation rate (for model)',
    'Service charge': 'Service charge (for model)',
}

# Create Document Number + Item key for matching
old_df['doc_item_key'] = old_df.apply(lambda x: f"{x['Document Number']}|{x['Item']}", axis=1)
new_df['doc_item_key'] = new_df.apply(lambda x: f"{x['Document Number']}|{x['Item']}", axis=1)

# Create lookup for old data
old_lookup = {}
for idx, row in old_df.iterrows():
    key = row['doc_item_key']
    if pd.notna(row['Document Number']) and pd.notna(row['Item']):
        old_lookup[key] = row

print(f'Old lookup created: {len(old_lookup)} Document Number + Item combinations')

# Analyze new file
new_records = 0
updated_records = 0
unchanged_records = 0
change_details = []

for idx, new_row in new_df.iterrows():
    key = new_row['doc_item_key'] 
    
    if pd.notna(new_row['Document Number']) and pd.notna(new_row['Item']):
        if key in old_lookup:
            # Existing combination - check for updates
            old_row = old_lookup[key]
            has_changes = False
            changed_columns = []
            
            for new_col, old_col in column_mapping.items():
                new_val = normalize_value(new_row.get(new_col))
                old_val = normalize_value(old_row.get(old_col))
                
                if new_val != old_val:
                    has_changes = True
                    changed_columns.append(new_col)
                    if len(change_details) < 10:  # Show first 10 examples
                        change_details.append({
                            'doc_number': new_row['Document Number'],
                            'item': new_row['Item'],
                            'column': new_col,
                            'old_value': old_val,
                            'new_value': new_val
                        })
            
            if has_changes:
                updated_records += 1
            else:
                unchanged_records += 1
        else:
            # New combination
            new_records += 1

print(f'\nANALYSIS RESULTS:')
print(f'New Document Number + Item combinations: {new_records}')
print(f'Updated existing combinations: {updated_records}') 
print(f'Unchanged combinations: {unchanged_records}')
print(f'Total processed: {new_records + updated_records + unchanged_records}')
print(f'Total new file rows: {len(new_df)}')

print(f'\nSample changes (first 10):')
for detail in change_details[:10]:
    print(f'  {detail["doc_number"]} | {detail["item"]}')
    print(f'    {detail["column"]}: "{detail["old_value"]}" -> "{detail["new_value"]}"')
    print()

print(f'\nDETAILED LOGIC:')
print(f'NEW ROWS (add new): Document Number + Item combinations that do NOT exist in old file')
print(f'UPDATE ROWS: Document Number + Item combinations that exist in old file BUT have changes in tracked columns')