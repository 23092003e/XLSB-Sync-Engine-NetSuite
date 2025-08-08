# scripts/process_entities.py
import os, glob, time, argparse
from excel_processor.config import DEFAULT_CONFIG
from excel_processor.batch import RobustBatchProcessor

def main():
    parser = argparse.ArgumentParser(description="Process XLSB entities with summary mapping")
    parser.add_argument("--entity-folder", required=True, help="Folder chứa các *.xlsb")
    parser.add_argument("--summary-path", required=True, help="Đường dẫn file tổng hợp Entities.xlsx")
    parser.add_argument("--mode", choices=["seq", "par"], default="seq",
                        help="seq=tuần tự (ổn định), par=‘song song bảo thủ’ (nhanh hơn)")
    args = parser.parse_args()

    file_paths = [fp for fp in glob.glob(os.path.join(args.entity_folder, "*.xlsb"))
                  if not os.path.basename(fp).startswith('~')]
    if not file_paths:
        print("❌ No XLSB files found!")
        return

    print(f"🎯 Found {len(file_paths)} files to process")
    print(f"📋 Files: {[os.path.basename(f) for f in file_paths]}")

    processor = RobustBatchProcessor(DEFAULT_CONFIG)
    t0 = time.time()
    if args.mode == "seq":
        print("\n🛡️ Using SEQUENTIAL mode")
        results = processor.process_files_sequential_robust(file_paths, args.summary_path)
    else:
        print("\n⚡ Using CONSERVATIVE PARALLEL mode")
        results = processor.process_files_parallel_conservative(file_paths, args.summary_path)
    total_time = time.time() - t0

    processor.print_enhanced_summary(results)
    log_file = os.path.join(args.entity_folder, "..", "processing_log.txt")
    try:
        with open(log_file, "w", encoding="utf-8") as f:
            f.write(f"Processing Results - {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*50 + "\n\n")
            for r in results:
                f.write(f"File: {os.path.basename(r.filepath)}\n"
                        f"Status: {r.status}\n"
                        f"Subsidiary: {r.subsidiary_found}\n"
                        f"Summary matches: {r.summary_matches}\n"
                        f"Rows updated: {r.rows_updated}\n"
                        f"Rows added: {r.rows_added}\n"
                        f"Processing time: {r.processing_time:.1f}s\n"
                        f"Error: {r.error_message}\n"
                        + "-"*30 + "\n")
        print(f"📄 Log saved to: {os.path.abspath(log_file)}")
    except Exception as e:
        print(f"⚠️ Could not save log: {e}")
    print(f"\n🏁 Total execution time: {total_time:.1f}s")

if __name__ == "__main__":
    main()
