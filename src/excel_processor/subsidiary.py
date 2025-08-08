# excel_processor/subsidiary.py
import os
import xlwings as xw

class SubsidiaryExtractor:
    @staticmethod
    def extract_subsidiary_enhanced(sheet: xw.Sheet, filepath: str, header_row: int = None) -> str:
        filename = os.path.basename(filepath)
        from_file = SubsidiaryExtractor._extract_from_filename(filename)
        if from_file:
            print(f"   🏢 Subsidiary from filename: {from_file}")
            return from_file

        if header_row:
            from_sheet = SubsidiaryExtractor._extract_from_sheet(sheet, header_row)
            if from_sheet:
                print(f"   🏢 Subsidiary from sheet: {from_sheet}")
                return from_sheet

        try:
            wb_name = sheet.book.name
            from_wb = SubsidiaryExtractor._extract_from_filename(wb_name)
            if from_wb:
                print(f"   🏢 Subsidiary from workbook: {from_wb}")
                return from_wb
        except:
            pass

        print(f"   ⚠️ Could not extract subsidiary from {filename}")
        return ""

    @staticmethod
    def _extract_from_filename(filename: str) -> str:
        name = filename.replace('.xlsb', '').replace('.xlsx', '')
        patterns = [
            lambda x: x.split('.')[1].strip().split('-')[0].strip() if '.' in x and len(x.split('.')) > 1 else None,
            lambda x: x.split('-')[0].strip() if '-' in x else None,
            lambda x: x.strip() if len(x.strip()) <= 5 and x.strip().isalpha() else None,
            lambda x: x[:5].strip() if x[:5].isupper() and len(x[:5].strip()) >= 3 else None
        ]
        for p in patterns:
            try:
                res = p(name)
                if res and len(res) >= 2:
                    return res.upper()
            except:
                continue
        return ""

    @staticmethod
    def _extract_from_sheet(sheet: xw.Sheet, header_row: int) -> str:
        try:
            for row in range(max(1, header_row - 3), header_row):
                for col in range(1, 6):
                    try:
                        v = sheet.range((row, col)).value
                        if v and isinstance(v, str):
                            if '-' in v and len(v.split('-')[0].strip()) <= 5:
                                return v.split('-')[0].strip().upper()
                    except:
                        continue
        except:
            pass
        return ""
