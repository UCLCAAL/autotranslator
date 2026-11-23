"""
Created on 19 Jun 2022

@author: turch
"""

import argparse
from set_parameters import Begin
import pathlib
from pathlib import Path
import logging
from logger_config import setup_logger
from encode import load_keys
from config import (
    DEFAULT_COLUMNS_ARCHIVE,
    DEFAULT_COLUMNS_MONUMENT,
    DEFAULT_START_ROW,
    FILE_TYPE_SHEET_MAP,
    SHEET_ARCHIVE,
    SHEET_MONUMENT
)

# Set up logger
logger = setup_logger()


def init() -> None:
    """
    Initialize and run the translation module.
    
    Parses command-line arguments and starts the translation process.
    """
    parser = argparse.ArgumentParser(
        description="Automatic translation module with glossary support"
    )
    current_file_location = pathlib.Path(__file__).parent.resolve()
    gloss_path = current_file_location / "glossary.xlsx"

    # Use current working directory as default instead of hardcoded Windows path
    default_data_path = Path.cwd() / "data"

    parser.add_argument(
        "-p",
        "--passphrase", 
        help="Team passphrase to unlock keys.enc", 
        default=None)
    
    parser.add_argument(
        "-d",
        "--data_path",
        type=str,
        default=str(default_data_path),
        help="Path to folder/file with data to translate",
    )
    parser.add_argument(
        "-il",
        "--input_language",
        type=str,
        choices=["ru", "en"],
        default="ru",
        help="Language of the input data (options: 'ru', 'en'), default is Russian (ru)",
    )
    parser.add_argument(
        "-ol",
        "--output_language",
        type=str,
        default="en",
        help="Language to translate the data into, default is English (en)",
    )
    parser.add_argument(
        "-g",
        "--glossfile",
        type=str,
        default="glossary.xlsx",
        help="Path to the glossary file (e.g., glossary.xlsx)",
    )

    parser.add_argument(
        "-s",
        "--sheet",
        type=str,
        choices=["Monument", "Archive", "other excel", "other filetype"],
        default="Archive",
        help="Sheet name to translate, default is Archives",
    )
    parser.add_argument(
        "-ic",
        "--input_columns",
        nargs="+",
        type=str,
        default="",
        help="The columns to be translated, default for 'Archives' is C, E, G, H, J; for 'Monuments' is B, C, D",
    )
    '''parser.add_argument(
        "-oc",
        "--output_columns",
        nargs="+",
        type=str,
        required=True,
        help="The columns to insert translated data into. Please type in corresponding order to input columns", 
    )''' # Removed as not needed for CSV output
    parser.add_argument(
        "-r",
        "--start_row",
        type=int,
        default=5,
        help="Row to start translation from (excluding column names, default: 5)",
    )

    parser.add_argument(
        "-od", 
        "--output_directory",
        type=Path,
        default=Path.cwd(),
        help="Path to CSV output directory",
    )

    args = parser.parse_args()

    api_keys = load_keys(args.passphrase)
    print(api_keys)
    
    # Determine sheet name and columns based on file type
    if args.sheet == "Monument":
        sheet = SHEET_MONUMENT
        columns = args.input_columns if args.input_columns else DEFAULT_COLUMNS_MONUMENT
    elif args.sheet == "Archive":
        sheet = SHEET_ARCHIVE
        columns = args.input_columns if args.input_columns else DEFAULT_COLUMNS_ARCHIVE
    else:
        sheet = input("Please enter name of the sheet to translate: ")
        columns = args.input_columns if args.input_columns else []

    # Validate glossary file
    gloss_file_path = Path(args.glossfile)
    if not gloss_file_path.is_absolute():
        gloss_file_path = current_file_location / args.glossfile
    
    if not gloss_file_path.exists():
        error_msg = f"Glossary file not found at {gloss_file_path}. Please provide a valid path."
        logger.error(error_msg)
        raise FileNotFoundError(error_msg)
    
    filetype = str(args.sheet)
    logger.info(f"Glossary path: {gloss_file_path}")
    logger.info(f"Columns to translate: {columns}")
    print(f"Glossary path: {gloss_file_path}")
    print(f"Columns: {columns}")
    
    try:
        begin = Begin()
        begin.modrun(
            data_path=args.data_path,
            ilanguage=args.input_language,
            olanguage=args.output_language,
            filetype=filetype,
            glossfile=str(gloss_file_path),
            sheet=sheet,
            input_column=columns,
            start_row=args.start_row,
            output_dir=args.output_directory
        )
    except Exception as e:
        logger.exception("Error during translation process")
        raise


init()
