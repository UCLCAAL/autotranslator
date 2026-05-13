"""
Utility functions for the autotranslator module.
"""
import logging
from pathlib import Path
from typing import List, Optional, Tuple
import pandas as pd

logger = logging.getLogger(__name__)


def validate_file_path(file_path: Path) -> bool:
    """
    Validate that a file path exists and is a file.
    
    Args:
        file_path: Path to validate
        
    Returns:
        True if valid, False otherwise
    """
    if not file_path.exists():
        logger.error(f"File does not exist: {file_path}")
        return False
    if not file_path.is_file():
        logger.error(f"Path is not a file: {file_path}")
        return False
    return True


def validate_directory(dir_path: Path) -> bool:
    """
    Validate that a directory path exists and is a directory.
    
    Args:
        dir_path: Path to validate
        
    Returns:
        True if valid, False otherwise
    """
    if not dir_path.exists():
        logger.error(f"Directory does not exist: {dir_path}")
        return False
    if not dir_path.is_dir():
        logger.error(f"Path is not a directory: {dir_path}")
        return False
    return True


def validate_excel_file(file_path: Path) -> Tuple[bool, Optional[str]]:
    """
    Validate that an Excel file can be opened and read.
    
    Args:
        file_path: Path to Excel file
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if not validate_file_path(file_path):
        return False, f"File does not exist or is not a file: {file_path}"
    
    try:
        # Try to read the file
        pd.read_excel(file_path, nrows=1)
        return True, None
    except Exception as e:
        error_msg = f"Cannot read Excel file {file_path}: {str(e)}"
        logger.error(error_msg)
        return False, error_msg


def validate_sheet_exists(file_path: Path, sheet_name: str) -> Tuple[bool, Optional[str]]:
    """
    Validate that a sheet exists in an Excel file.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Name of the sheet to check
        
    Returns:
        Tuple of (exists, error_message)
    """
    try:
        excel_file = pd.ExcelFile(file_path)
        if sheet_name not in excel_file.sheet_names:
            error_msg = f"Sheet '{sheet_name}' not found in {file_path}. Available sheets: {excel_file.sheet_names}"
            logger.error(error_msg)
            return False, error_msg
        return True, None
    except Exception as e:
        error_msg = f"Error reading Excel file {file_path}: {str(e)}"
        logger.error(error_msg)
        return False, error_msg


def validate_columns_exist(file_path: Path, sheet_name: str, columns: List[str]) -> Tuple[bool, Optional[str]]:
    """
    Validate that specified columns exist in the Excel sheet.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Name of the sheet
        columns: List of column letters/names to validate
        
    Returns:
        Tuple of (all_exist, error_message)
    """
    try:
        # Read header row to get available columns
        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0)
        available_columns = list(df.columns)
        
        # For column letters, we need to check if they can be accessed
        # This is a simplified check - in practice, Excel column letters need conversion
        # For now, we'll just try to read them
        try:
            test_df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns, nrows=1)
            return True, None
        except Exception as e:
            error_msg = f"Columns {columns} cannot be accessed in sheet '{sheet_name}': {str(e)}"
            logger.error(error_msg)
            return False, error_msg
    except Exception as e:
        error_msg = f"Error validating columns in {file_path}: {str(e)}"
        logger.error(error_msg)
        return False, error_msg


def ensure_output_directory(output_dir: Path, create_subdir: bool = False, subdir_name: Optional[str] = None) -> Path:
    """
    Ensure output directory exists, optionally creating a subdirectory.
    
    Args:
        output_dir: Base output directory
        create_subdir: Whether to create a subdirectory
        subdir_name: Name of subdirectory (if None, uses default)
        
    Returns:
        Path to the output directory (or subdirectory if created)
    """
    if create_subdir and subdir_name:
        output_path = output_dir / subdir_name
    else:
        output_path = output_dir
    
    output_path.mkdir(parents=True, exist_ok=True)
    logger.debug(f"Output directory ensured: {output_path}")
    return output_path
