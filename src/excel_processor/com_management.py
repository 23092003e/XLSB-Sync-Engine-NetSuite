# excel_processor/com_management.py
import time, gc, subprocess
import xlwings as xw
import pythoncom

class COMManager:
    @staticmethod
    def initialize_com() -> bool:
        try:
            pythoncom.CoInitialize()
            return True
        except Exception as e:
            print(f"COM initialization failed: {e}")
            return False

    @staticmethod
    def cleanup_com():
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"COM cleanup warning: {e}")

    @staticmethod
    def kill_excel_processes():
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'excel.exe'],
                           capture_output=True, timeout=5)
        except Exception as e:
            print(f"Excel cleanup warning: {e}")

class EnhancedExcelOptimizer:
    @staticmethod
    def setup_excel_app_robust():
        app = None
        for attempt in range(3):
            try:
                if not COMManager.initialize_com():
                    continue
                app = xw.App(visible=False, add_book=False)
                time.sleep(0.2)  # Reduced wait time
                _ = app.version  # test
                # Set Excel properties individually with error handling
                try:
                    app.screen_updating = False
                except Exception as e:
                    print(f"Warning: screen_updating failed: {e}")
                
                try:
                    app.display_alerts = False
                except Exception as e:
                    print(f"Warning: display_alerts failed: {e}")
                
                try:
                    app.enable_events = False
                except Exception as e:
                    print(f"Warning: enable_events failed: {e}")
                
                try:
                    app.calculation = 'manual'
                except Exception as e:
                    print(f"Warning: calculation setting failed: {e}")
                    # Try alternative approach
                    try:
                        app.calculation = -4135  # xlCalculationManual constant
                    except Exception as e2:
                        print(f"Warning: alternative calculation setting failed: {e2}")
                
                try:
                    app.interactive = False
                except Exception as e:
                    print(f"Warning: interactive setting failed: {e}")
                print(f"   📱 Excel initialized with optimizations (attempt {attempt+1})")
                return app
            except Exception as e:
                print(f"   ⚠️ Excel setup attempt {attempt+1} failed: {e}")
                try:
                    if app: app.quit()
                except: pass
                app = None
                time.sleep((attempt+1) * 0.2)  # Reduced retry delay
        return None

    @staticmethod
    def safe_excel_operation(func, *args, **kwargs):
        for retry in range(2):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if retry == 1:
                    raise e
                time.sleep(0.05)  # Reduced retry delay
                gc.collect()

    @staticmethod
    def find_header_row_enhanced(sheet: xw.Sheet):
        for r in range(1, 8):
            try:
                vals = EnhancedExcelOptimizer.safe_excel_operation(
                    lambda: sheet.range((r, 1), (r, 15)).value
                )
                if vals:
                    vals_str = ' '.join(str(v) for v in vals if v)
                    if 'Item2' in vals_str and 'Note' in vals_str:
                        print(f"   📍 Header at row {r}")
                        return r
            except Exception as e:
                print(f"   ⚠️ check row {r}: {e}")
                continue
        return None
