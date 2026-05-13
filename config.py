"""
Configuration constants for the autotranslator module.
"""
from pathlib import Path
from typing import Dict, List

# Default sheet names
SHEET_MONUMENT = "Data Sheet"
SHEET_ARCHIVE = "2.Описание"

# Default column mappings
DEFAULT_COLUMNS_ARCHIVE = ['C', 'E', 'G', 'H', 'J']
DEFAULT_COLUMNS_MONUMENT = ['B', 'C', 'D']

# Default start row (excluding column headers)
DEFAULT_START_ROW = 5

# Row offset for max_row calculation
MAX_ROW_OFFSET = 3

# Columns used for counting rows (CAAL_ID columns)
ROW_COUNT_COLUMN_ARCHIVE = 'C'
ROW_COUNT_COLUMN_MONUMENT = 'G'

# Columns to combine in output (original + translated)
COMBINE_COLUMNS = ['Name', 'Description']

# Output CSV configuration
CSV_DELIMITER = ';'
CSV_ENCODING = 'utf-8-sig'

# Column name mappings for translation
COLUMN_NAME_MAPPINGS: Dict[str, str] = {
    'name': 'Name',
    'quantity': 'Quantity',
    'caal': 'CAAL_ID',
    'passport': 'Description'
}

# Supported languages
SUPPORTED_INPUT_LANGUAGES = ["ru", "en"]
SUPPORTED_OUTPUT_LANGUAGES = ["en", "ru"]

# File type mappings
FILE_TYPE_SHEET_MAP: Dict[str, str] = {
    "Monument": SHEET_MONUMENT,
    "Archive": SHEET_ARCHIVE
}

# Default output subdirectory (if needed)
OUTPUT_SUBDIR = "translated_csv"
