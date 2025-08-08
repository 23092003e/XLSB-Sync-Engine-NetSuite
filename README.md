# XLSB-Sync-Engine-NetSuite
An automated data integration tool designed to streamline the process of updating financial models stored in XLSB format across multiple subsidiaries

# Excel Processor (refactor)
## How to Run
```bash
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt

python scripts/process_entities.py ^
  --entity-folder "C:\Users\ADMIN\Desktop\Process XLSB\data\input\entities" ^
  --summary-path  "C:\Users\ADMIN\Desktop\Process XLSB\data\input\summary\File tổng hợp Entities.xlsx" ^
  --mode seq
