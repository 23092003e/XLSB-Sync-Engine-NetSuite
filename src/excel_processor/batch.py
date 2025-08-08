# excel_processor/batch.py
import os, time, gc
from typing import List
from concurrent.futures import ThreadPoolExecutor, as_completed

from .models import ProcessingConfig, ProcessingResult
from .com_management import COMManager
from .processor import EnhancedExcelProcessor

class RobustBatchProcessor:
    def __init__(self, config: ProcessingConfig):
        self.config = config

    def process_files_sequential_robust(self, file_paths: List[str], summary_path: str) -> List[ProcessingResult]:
        print(f"🚀 Starting SEQUENTIAL ROBUST processing")
        print(f"   📁 Files: {len(file_paths)} | 🛡️ Mode: Sequential")

        COMManager.kill_excel_processes(); time.sleep(2)
        processor = EnhancedExcelProcessor(self.config)
        processor.load_summary_data_enhanced(summary_path)

        results = []
        for i, fp in enumerate(file_paths):
            print(f"\n📦 Processing file {i+1}/{len(file_paths)}: {os.path.basename(fp)}")
            result = None
            for attempt in range(self.config.retry_attempts):
                if attempt > 0:
                    print(f"   🔄 Retry attempt {attempt+1}")
                    time.sleep(2)
                result = processor.process_single_file_enhanced(fp)
                if result.status == 'success':
                    break
                COMManager.kill_excel_processes(); time.sleep(1)
            results.append(result)
            gc.collect(); time.sleep(0.1)  # Reduced delay between files
        return results

    def process_files_parallel_conservative(self, file_paths: List[str], summary_path: str) -> List[ProcessingResult]:
        print(f"🚀 Starting CONSERVATIVE PARALLEL processing")
        print(f"   📁 Files: {len(file_paths)} | 🔧 Max workers: {self.config.max_excel_instances}")

        COMManager.kill_excel_processes(); time.sleep(2)
        processor = EnhancedExcelProcessor(self.config)
        processor.load_summary_data_enhanced(summary_path)

        batch_size = self.config.max_excel_instances
        batches = [file_paths[i:i+batch_size] for i in range(0, len(file_paths), batch_size)]
        all_results = []
        for bi, batch in enumerate(batches):
            print(f"\n📦 Processing batch {bi+1}/{len(batches)}")
            with ThreadPoolExecutor(max_workers=len(batch)) as ex:
                fut_map = {ex.submit(self._process_with_retry, processor, fp): fp for fp in batch}
                batch_results = []
                for fut in as_completed(fut_map, timeout=self.config.timeout_seconds):
                    fp = fut_map[fut]
                    try:
                        res = fut.result()
                        batch_results.append(res)
                        if res.status == 'success':
                            print(f"   ✅ {os.path.basename(fp)}: {res.rows_updated} upd, {res.rows_added} add")
                        else:
                            print(f"   ❌ {os.path.basename(fp)}: {res.error_message}")
                    except Exception as e:
                        print(f"   💥 {os.path.basename(fp)} failed: {e}")
                        batch_results.append(ProcessingResult(filepath=fp, status='error',
                                                              error_message=f"Execution failed: {e}"))
            all_results.extend(batch_results)
            if bi < len(batches)-1:
                print("   🧹 Cleaning up between batches...")
                COMManager.kill_excel_processes(); gc.collect(); time.sleep(1)  # Reduced cleanup delay
        return all_results

    def _process_with_retry(self, processor: EnhancedExcelProcessor, filepath: str) -> ProcessingResult:
        res = None
        for attempt in range(self.config.retry_attempts):
            if attempt > 0:
                print(f"   🔄 Retrying {os.path.basename(filepath)} (attempt {attempt+1})")
                time.sleep(0.5)  # Reduced retry delay
            res = processor.process_single_file_enhanced(filepath)
            if res.status == 'success':
                return res
            COMManager.kill_excel_processes(); time.sleep(0.2)  # Reduced process kill delay
        return res

    def print_enhanced_summary(self, results: List[ProcessingResult]):
        ok = [r for r in results if r.status == 'success']
        bad = [r for r in results if r.status == 'error']
        total_time = sum(r.processing_time for r in ok)
        total_updated = sum(r.rows_updated for r in ok)
        total_added = sum(r.rows_added for r in ok)
        print("\n📊 ENHANCED PROCESSING SUMMARY")
        print(f"   ✅ Successful: {len(ok)}/{len(results)}")
        print(f"   ❌ Failed: {len(bad)}")
        print(f"   📝 Total updated rows: {total_updated}")
        print(f"   ➕ Total added rows: {total_added}")
        if ok:
            print(f"   ⏱️ Avg time/file: {total_time/len(ok):.1f}s")
