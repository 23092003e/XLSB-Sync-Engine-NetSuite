# excel_processor/com_management.py
import time, gc, subprocess
import xlwings as xw
import pythoncom

class COMManager:
    _com_initialized = False
    _active_apps = {}  # Pool of Excel applications
    _app_usage_count = {}  # Track usage for cleanup
    
    @staticmethod
    def initialize_com() -> bool:
        try:
            if not COMManager._com_initialized:
                pythoncom.CoInitialize()
                COMManager._com_initialized = True
            return True
        except Exception as e:
            print(f"COM initialization failed: {e}")
            return False

    @staticmethod
    def cleanup_com():
        try:
            # Clean up all pooled applications first
            COMManager.cleanup_all_excel_apps()
            
            if COMManager._com_initialized:
                pythoncom.CoUninitialize()
                COMManager._com_initialized = False
        except Exception as e:
            print(f"COM cleanup warning: {e}")

    @staticmethod
    def get_or_create_excel_app(app_id: str = None) -> xw.App:
        """Get existing Excel app from pool or create new one"""
        if app_id is None:
            app_id = f"app_{len(COMManager._active_apps)}"
            
        if app_id in COMManager._active_apps:
            app = COMManager._active_apps[app_id]
            try:
                # Test if app is still alive
                _ = app.version
                COMManager._app_usage_count[app_id] = COMManager._app_usage_count.get(app_id, 0) + 1
                return app
            except:
                # App is dead, remove from pool
                COMManager._active_apps.pop(app_id, None)
                COMManager._app_usage_count.pop(app_id, None)
        
        # Create new app
        app = EnhancedExcelOptimizer.setup_excel_app_robust()
        if app:
            COMManager._active_apps[app_id] = app
            COMManager._app_usage_count[app_id] = 1
            print(f"   [POOL] Created new Excel app: {app_id} (Pool size: {len(COMManager._active_apps)})")
        
        return app

    @staticmethod
    def release_excel_app(app_id: str):
        """Release Excel app back to pool (or close if overused)"""
        if app_id not in COMManager._active_apps:
            return
            
        usage_count = COMManager._app_usage_count.get(app_id, 0)
        
        # Close app if it's been used too many times to prevent memory leaks
        if usage_count > 50:
            try:
                app = COMManager._active_apps[app_id]
                app.quit()
                print(f"   [POOL] Closed overused Excel app: {app_id} (used {usage_count} times)")
            except Exception as e:
                print(f"   [WARNING] Error closing Excel app {app_id}: {e}")
            finally:
                COMManager._active_apps.pop(app_id, None)
                COMManager._app_usage_count.pop(app_id, None)

    @staticmethod
    def cleanup_all_excel_apps():
        """Clean up all Excel applications in the pool"""
        for app_id, app in list(COMManager._active_apps.items()):
            try:
                app.quit()
                print(f"   [CLEANUP] Cleaned up Excel app: {app_id}")
            except Exception as e:
                print(f"   [WARNING] Error cleaning up Excel app {app_id}: {e}")
        
        COMManager._active_apps.clear()
        COMManager._app_usage_count.clear()

    @staticmethod
    def kill_excel_processes():
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'excel.exe'],
                           capture_output=True, timeout=5)
        except Exception as e:
            print(f"Excel cleanup warning: {e}")

    @staticmethod
    def get_pool_stats():
        """Get statistics about the Excel application pool"""
        return {
            'active_apps': len(COMManager._active_apps),
            'total_usage': sum(COMManager._app_usage_count.values()),
            'apps': dict(COMManager._app_usage_count)
        }

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
                print(f"   [EXCEL] Excel initialized with optimizations (attempt {attempt+1})")
                return app
            except Exception as e:
                print(f"   [WARNING] Excel setup attempt {attempt+1} failed: {e}")
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
