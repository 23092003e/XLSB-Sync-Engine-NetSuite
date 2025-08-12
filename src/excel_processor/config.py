# excel_processor/config.py
from .models import ProcessingConfig

# Mapping for column names in the Excel files

COLUMN_MAPPING = {
    'Unit name': 'Factory code',
    'Tenant ID': 'Tenant code',
    'Tenant': 'Tenant name',
    'GLA': 'GLA',
    'Contract type': 'Existing/New/Exp/Renew',
    'Rent USD_Item (for model)': 'Rent (USD)',
    'Rent VND_Item (for model)': 'Rent (VND)',
    'Total months fitout & rent free (for model)': 'Rent free',
    'Service charge (for model)': 'Service charge',
    'Escalation rate (for model)': 'Growth rate (Act)',
    'Broker? (Yes/No)': 'Broker',
    'End date (for model)': 'End date',
    'Start date (for model)': 'Start date',
    'UFL Status': 'Handover',
    'Payment term (for model)': 'Payment term',
    'Months fit-out (for model)': 'Fitting out'
}

def get_optimal_worker_count() -> int:
    """Calculate optimal number of Excel workers based on system resources"""
    import os
    import psutil
    
    # Get CPU cores (physical cores preferred)
    cpu_cores = psutil.cpu_count(logical=False) or psutil.cpu_count()
    
    # Get available memory in GB
    memory_gb = psutil.virtual_memory().total / (1024**3)
    
    # FIXED: Conservative approach for Excel COM objects
    # Excel instances are heavy and don't parallelize well with threading
    # Use 1 worker per 2 cores, maximum 4 workers for stability
    max_workers_by_cpu = min(max(1, cpu_cores // 2), 4)  # Conservative: max 4 workers
    max_workers_by_memory = max(1, int(memory_gb / 2))  # 2GB per worker (more realistic)
    
    optimal_workers = min(max_workers_by_cpu, max_workers_by_memory)
    
    print(f"   System: {cpu_cores} cores, {memory_gb:.1f}GB RAM")
    print(f"   Optimal workers: {optimal_workers} (CPU: {max_workers_by_cpu}, Memory: {max_workers_by_memory})")
    
    return optimal_workers

# Dynamic configuration based on system capabilities
optimal_workers = get_optimal_worker_count()

DEFAULT_CONFIG = ProcessingConfig(
    max_excel_instances=optimal_workers,
    timeout_seconds=900,      # Increased timeout for large files (15 minutes)
    retry_attempts=2,         # Balanced retry attempts
    backup_enabled=True,
    column_mapping=COLUMN_MAPPING,
    # Enable all performance optimizations
    auto_detect_optimal_workers=True,
    enable_memory_monitoring=True,
    enable_chunked_processing=True,
    chunk_size=5000,          # Optimized for large files
    memory_threshold_percent=70.0,  # Conservative memory threshold
)
