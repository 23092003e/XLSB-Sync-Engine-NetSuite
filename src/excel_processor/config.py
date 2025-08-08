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
    'Contract status': 'Handover'
}

DEFAULT_CONFIG = ProcessingConfig(
    max_excel_instances=2,
    timeout_seconds=300,
    retry_attempts=2,
    backup_enabled=True,
    column_mapping=COLUMN_MAPPING
)
