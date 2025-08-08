# excel_processor/models.py
from dataclasses import dataclass, field
from typing import Dict

@dataclass
class ProcessingConfig:
    max_excel_instances: int = 2
    chunk_size: int = 1000
    memory_threshold_percent: float = 80.0
    timeout_seconds: int = 300
    backup_enabled: bool = True
    retry_attempts: int = 2
    excel_startup_delay: float = 1.0
    column_mapping: Dict[str, str] = field(default_factory=dict)

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
