# excel_processor/models.py
from dataclasses import dataclass, field
from typing import Dict

@dataclass
class ProcessingConfig:
    max_excel_instances: int = 2
    chunk_size: int = 5000  # Increased default chunk size for large files
    memory_threshold_percent: float = 70.0  # Reduced from 80% for better safety
    timeout_seconds: int = 900  # Increased for large files (15 minutes)
    backup_enabled: bool = True
    retry_attempts: int = 2
    excel_startup_delay: float = 1.0
    column_mapping: Dict[str, str] = field(default_factory=dict)
    
    # New configuration options for performance optimization
    auto_detect_optimal_workers: bool = True
    max_workers_override: int = None  # Override auto-detection if set
    enable_memory_monitoring: bool = True
    enable_chunked_processing: bool = True
    min_chunk_size: int = 500
    max_chunk_size: int = 10000
    memory_cleanup_interval: int = 5  # Cleanup every N chunks

@dataclass
class ProcessingResult:
    filepath: str
    status: str  # 'success', 'error', 'skipped'
    rows_updated: int = 0
    rows_added: int = 0
    processing_time: float = 0.0
    memory_used_mb: float = 0.0
    error_message: str = ""
    subsidiary_found: str = ""
    summary_matches: int = 0
