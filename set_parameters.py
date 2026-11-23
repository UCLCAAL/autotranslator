"""
Created on 18 Dec 2021

@author: turch
"""

from translate_excel_to_csv import TranslateRun
import pandas as pd
from pathlib import Path
import logging
from typing import List, Optional
from config import ROW_COUNT_COLUMN_ARCHIVE, ROW_COUNT_COLUMN_MONUMENT, MAX_ROW_OFFSET
from utils import validate_excel_file, validate_sheet_exists

logger = logging.getLogger(__name__)


class LoadFolders:
    """
    Class to load and organize folders/files for translation.

    Handles discovery and organization of Excel files from specified directories,
    separating them into Archive and Monument categories.
    """

    def __init__(self) -> None:
        """Initialize LoadFolders with instance-level lists to avoid shared state."""
        self.archive: List[str] = []  # list for archive spreadsheets
        self.monument: List[str] = []  # list for monument spreadsheets
        self.docs: List[str] = []  # list for document files
        self.biglist: List[List[str]] = [
            self.archive,
            self.monument,
        ]  # combined list for iteration
        self.l: Optional[List[str]] = None  # current list being processed

    def sortfolder(self, data_path: Path, filetype: str, folder: List[str]) -> None:
        """
        Sort files in a folder, adding Excel files to the specified folder list.

        Args:
            data_path: Path to directory or file to process
            filetype: Type of file (not currently used but kept for compatibility)
            folder: List to append file paths to
        """
        p = Path(data_path)
        if ".xl" in p.name.lower():
            folder.append(str(p))
        else:
            for file_path in p.rglob("*"):
                if file_path.is_file() and ".xl" in file_path.suffix.lower():
                    folder.append(str(file_path))

    def openfolder(self, data_path: str, filetype: str) -> None:
        """
        Open and organize folders/files from the specified data path.

        This function locates folders from the specified data path, accesses files within them,
        and extracts file paths to specified lists. File paths are then used to access
        individual files for translation.

        Args:
            data_path: Path to directory or file containing data to translate
            filetype: Type of files to process (Archive, Monument, Excel, etc.)
        """
        p = Path(str(data_path).replace('"', ""))

        if filetype in ["Archive", "Monument", "Excel"]:
            if "archive" in p.name.lower():
                folder = self.archive
                self.sortfolder(p, filetype, folder)
            elif "monument" in p.name.lower():
                folder = self.monument
                self.sortfolder(p, filetype, folder)
            else:
                for subdir in p.rglob("*"):
                    # print('subdir: ' + str(subdir)) #test that this is working
                    if subdir.is_dir():
                        if "archive" in subdir.name.lower():
                            folder = self.archive
                            self.sortfolder(subdir, filetype, folder)
                        elif "monument" in subdir.name.lower():
                            folder = self.monument
                            self.sortfolder(subdir, filetype, folder)
                        else:
                            pass
        else:
            if p.suffix in [".docx", ".pdf"]:
                self.docs.append(str(p))
                self.dirpath = str(p.parent)
            else:
                # Walk subdirectories and files
                if any(subdir.is_dir() for subdir in p.iterdir()):
                    for subdir in p.iterdir():
                        if subdir.is_dir():
                            self.dirpath = str(subdir)
                            for file_path in subdir.rglob("*"):
                                if file_path.is_file():
                                    self.docs.append(str(file_path))
                else:
                    for file_path in p.iterdir():
                        if file_path.is_file():
                            self.dirpath = str(p)
                            self.docs.append(str(file_path))

    def runtransl(
        self,
        ilanguage: str,
        olanguage: str,
        data_path: str,
        filetype: str,
        glossfile: str,
        sheet: str,
        input_column: List[str],
        start_row: int,
        output_dir: Path,
        api_keys: List[str],
    ) -> None:
        """
        Run translation workflow for discovered files.

        Args:
            ilanguage: Input language code
            olanguage: Output language code
            data_path: Path to data directory
            filetype: Type of files (Archive, Monument, Excel, etc.)
            glossfile: Path to glossary file
            sheet: Sheet name to translate
            input_column: List of column letters to translate
            start_row: Starting row number
            output_dir: Output directory for CSV files
        """

        def show_folders(file_list: List[str], folder_type: str) -> None:
            """Display list of files to be translated."""
            logger.info(f"Folder length: {len(file_list)}")
            self.l = file_list
            logger.info(
                f"The folder being translated is {folder_type}. The files in this folder are:"
            )
            print(f"length {len(file_list)}")
            print(
                f"The folder being translated is {folder_type}. The files in this folder are:"
            )
            for idx, f in enumerate(file_list, start=1):
                name = f"{idx}. {f}"
                print(name)

        if len(self.archive) == 0 and len(self.monument) == 0 and len(self.docs) == 0:
            error_msg = "Empty input! Did you select the wrong filetype or put in an invalid path? Reset and try again."
            logger.warning(error_msg)
            print(error_msg)

        if filetype in ["Excel", "Archive", "Monument"]:
            for file_list in self.biglist:
                while True:
                    if file_list == self.archive:
                        folder_type = "Archives"
                        if len(file_list) == 0:
                            break
                        else:
                            show_folders(file_list, folder_type)

                    elif file_list == self.monument:
                        folder_type = "Monuments"
                        if len(file_list) == 0:
                            break
                        else:
                            show_folders(file_list, folder_type)
                    else:
                        continue

                    remy = str(
                        input(
                            "\n If you want to remove any spreadsheets, enter the corresponding numbers "
                            "here separated by comma, otherwise enter N:"
                        )
                    )
                    if remy:
                        self.removefiles(
                            remy,
                            file_list,
                            data_path,
                            ilanguage,
                            olanguage,
                            filetype,
                            glossfile,
                            sheet,
                            input_column,
                            start_row,
                            output_dir,
                            api_keys,
                        )
                        break

        else:  # if other types of files - PDF, DOCS
            file_list = self.docs
            self.l = file_list
            logger.info("The files to be translated are:")
            print("The files to be translated are:")

            for idx, f in enumerate(file_list, start=1):
                name = f"{idx}. {f}"
                print(name)

            remy = input(
                "If you want to remove files, enter corresponding numbers here separated by comma, "
                "otherwise enter N:"
            )
            if remy:
                self.removefiles(
                    remy,
                    file_list,
                    data_path,
                    ilanguage,
                    olanguage,
                    filetype,
                    glossfile,
                    sheet,
                    input_column,
                    start_row,
                    output_dir,
                    api_keys,
                )

    def removefiles(
        self,
        remy: str,
        file_list: List[str],
        data_path: str,
        ilanguage: str,
        olanguage: str,
        filetype: str,
        glossfile: str,
        sheet: str,
        input_column: List[str],
        start_row: int,
        output_dir: Path,
        api_keys: List[str],
    ) -> None:
        """
        Allow user to remove files from the translation list by index.

        This is useful for files that have already been translated, have problems,
        or should be avoided for any reason.

        Args:
            remy: User input string with file indices to remove (comma-separated) or 'N'
            file_list: List of file paths
            data_path: Path to data directory
            ilanguage: Input language code
            olanguage: Output language code
            filetype: Type of files
            glossfile: Path to glossary file
            sheet: Sheet name to translate
            input_column: List of column letters to translate
            start_row: Starting row number
            output_dir: Output directory for CSV files
            api_keys: List of API keys
        """
        if str(remy) == "N":  # if No files are to be removed
            self.xl_proceed(
                sheet,
                start_row,
                ilanguage,
                olanguage,
                input_column,
                filetype,
                glossfile,
                data_path,
                output_dir,
                api_keys,
            )
        elif str(remy) == "":  # if it's a blank (clicked OK by accident)
            self.runtransl(
                ilanguage,
                olanguage,
                data_path,
                filetype,
                glossfile,
                sheet,
                input_column,
                start_row,
                output_dir,
                api_keys,
            )
        else:  # otherwise
            if "," in str(remy):  # if comma is in the string
                rem = [r.strip() for r in remy.split(",")]  # split and strip values
                # Sort indices in reverse order to avoid index shifting issues
                indices_to_remove = sorted(
                    [int(r) - 1 for r in rem if r.isdigit()], reverse=True
                )
                for idx in indices_to_remove:
                    if 0 <= idx < len(file_list):
                        del file_list[idx]
                logger.info(f"Removed {len(indices_to_remove)} files from list")
                self.runtransl(
                    ilanguage,
                    olanguage,
                    data_path,
                    filetype,
                    glossfile,
                    sheet,
                    input_column,
                    start_row,
                    output_dir,
                    api_keys,
                )
            elif str(remy).isdigit():
                idx = int(remy) - 1
                if 0 <= idx < len(file_list):
                    del file_list[idx]
                    logger.info(f"Removed file at index {idx}")
                self.runtransl(
                    ilanguage,
                    olanguage,
                    data_path,
                    filetype,
                    glossfile,
                    sheet,
                    input_column,
                    start_row,
                    output_dir,
                    api_keys,
                )
            else:
                error_msg = "Invalid Input"
                logger.warning(error_msg)
                print(error_msg)
                self.runtransl(
                    ilanguage,
                    olanguage,
                    data_path,
                    filetype,
                    glossfile,
                    sheet,
                    input_column,
                    start_row,
                    output_dir,
                    api_keys,
                )

        """
        
            Here the user is able to put in which sheet the program should access and which columns they would like to translate and in what order. This is left up to the user so that if only a particular 
            column needs to be translated quickly for multiple files, that can be accounted for as well as, having all the needed fields be translated by putting in multiple columns. The format and order are critical here:
            the format because there are columns like AK or BD made up of 2 characters that still need to be read as one entry, but there is also a need to combine data from 2 Russian-entry columns into 1 English-entry
            column. As such if a / is placed then the columns are translated together and output as 1 entry; if a comma is placed, then that means the columns should be translated separately with separate outputs. T
            The order has to be the same for the input and output as that is the only way the program knows which column corresponds to which
            
        """

    # table = self.tables[0] #once each file is cycled through the SQL table can be accessed for analysis - in this case the Archive table is used but either can be depending on what
    # is wanted for the analysis
    # SQLite.analysis(SQLite, table, self.SQLcolumns) #call SQLite analysis to find number of different monuments

    def edit_inputs(
        self,
        input_sheet: str,
        start_row: int,
        ilanguage: str,
        olanguage: str,
        input_column: List[str],
        filetype: str,
        glossfile: str,
        data_path: str,
        output_dir: Path,
        api_keys: List[str],
    ) -> None:
        """
        Allow user to edit input parameters interactively.

        Args:
            input_sheet: Current sheet name
            start_row: Current start row
            ilanguage: Input language code
            olanguage: Output language code
            input_column: Current column list
            filetype: Type of files
            glossfile: Path to glossary file
            data_path: Path to data directory
            output_dir: Output directory for CSV files
            api_keys: List of API keys
        """
        from config import SHEET_MONUMENT, SHEET_ARCHIVE

        sheet_in = input(
            "Enter 'arg' to keep same input, or enter sheet in spreadsheet (ex: 'Data Sheet'):"
        )
        input_col_in = input(
            'Enter "arg" to keep same input or enter columns to be translated separated by space'
        )
        start_row_in = input(
            'Enter "arg" to keep same input or enter row number from which to start translation '
            "(ex: 5 - exclude column names):"
        )

        if sheet_in.lower() != "arg":
            input_sheet = sheet_in

        if input_col_in.lower() != "arg":
            input_column = input_col_in.split()

        if start_row_in.lower() != "arg":
            try:
                start_row = int(start_row_in)
            except ValueError:
                logger.warning(
                    f"Invalid start_row input: {start_row_in}. Using default."
                )
                print(f"Invalid start_row input: {start_row_in}. Using default.")

        if input_sheet == "Monuments":
            sheet = SHEET_MONUMENT
        elif input_sheet == "Archive":
            sheet = SHEET_ARCHIVE
        else:
            sheet = input_sheet

        self.xl_proceed(
            sheet,
            start_row,
            ilanguage,
            olanguage,
            input_column,
            filetype,
            glossfile,
            data_path,
            output_dir,
            api_keys,
        )

    def xl_proceed(
        self,
        input_sheet: str,
        start_row: int,
        ilanguage: str,
        olanguage: str,
        input_column: List[str],
        filetype: str,
        glossfile: str,
        data_path: str,
        output_dir: Path,
        api_keys: List[str],
    ) -> None:
        """
        Proceed with Excel file translation after user confirmation.

        Args:
            input_sheet: Sheet name to translate
            start_row: Starting row number
            ilanguage: Input language code
            olanguage: Output language code
            input_column: List of column letters to translate
            filetype: Type of files
            glossfile: Path to glossary file
            data_path: Path to data directory
            output_dir: Output directory for CSV files
            api_keys: List of API keys
        """
        transl = TranslateRun()

        # Display configuration
        txt1 = f"Sheet: {input_sheet}\nColumns to translate:\n"
        for col in input_column:
            txt1 += f"{col}\n"
        print(f"{txt1}\nStarting row: {start_row}")
        logger.info(
            f"Translation configuration - Sheet: {input_sheet}, Columns: {input_column}, Start row: {start_row}"
        )

        def process_translation_batch() -> None:
            """
            Process translation for all files in the list.

            Handles errors gracefully by continuing with remaining files
            and providing a summary at the end.
            """
            if not self.l:
                logger.warning("No files to process")
                print("No files to process")
                return

            successful_files = []
            failed_files = []

            for file in self.l:
                logger.info(f"Processing file: {file}")
                print(f"\nProcessing file: {file}")

                try:
                    # Validate file before processing
                    file_path = Path(file)
                    is_valid, error_msg = validate_excel_file(file_path)
                    if not is_valid:
                        logger.error(f"Skipping invalid file {file}: {error_msg}")
                        print(f"  ERROR: Skipping - {error_msg}")
                        failed_files.append((file, error_msg))
                        continue

                    # Validate sheet exists
                    sheet_exists, sheet_error = validate_sheet_exists(
                        file_path, input_sheet
                    )
                    if not sheet_exists:
                        logger.error(f"Skipping {file}: {sheet_error}")
                        print(f"  ERROR: Skipping - {sheet_error}")
                        failed_files.append((file, sheet_error))
                        continue

                    # Determine column for row counting based on file type
                    if file in self.archive:
                        cols = ROW_COUNT_COLUMN_ARCHIVE
                    else:
                        cols = ROW_COUNT_COLUMN_MONUMENT

                    # Read column to determine max_row (optimized: read once)
                    row_length = pd.read_excel(file, input_sheet, usecols=cols)
                    max_row = int(len(row_length.dropna())) + MAX_ROW_OFFSET

                    if len(row_length) < start_row:  # skip empty files
                        warning_msg = (
                            f"insufficient rows ({len(row_length)} < {start_row})"
                        )
                        logger.warning(f"Skipping {file}: {warning_msg}")
                        print(f"  WARNING: Skipping - {warning_msg}")
                        failed_files.append((file, warning_msg))
                        continue

                    # Process translation
                    transl.runtransl(
                        file,
                        input_sheet,
                        input_column,
                        start_row,
                        max_row,
                        ilanguage,
                        olanguage,
                        filetype,
                        glossfile,
                        output_dir,
                        api_keys,
                    )
                    successful_files.append(file)
                    print("  ✓ Successfully translated")

                except PermissionError as e:
                    error_msg = (
                        f"Permission denied. Please check file permissions and ensure "
                        f"the file is not open in another application: {str(e)}"
                    )
                    logger.error(f"Permission error for {file}: {str(e)}")
                    print(f"  ERROR: {error_msg}")
                    failed_files.append((file, error_msg))
                    # Continue with next file instead of stopping entire batch
                    continue

                except Exception as e:
                    error_msg = f"Unexpected error: {str(e)}"
                    logger.exception(f"Error processing {file}")
                    print(f"  ERROR: {error_msg}")
                    failed_files.append((file, error_msg))
                    # Continue with next file instead of stopping entire batch
                    continue

            # Print summary
            print(f"\n{'=' * 60}")
            print("Translation Summary:")
            print(f"  Successful: {len(successful_files)}")
            print(f"  Failed: {len(failed_files)}")
            print(f"{'=' * 60}")

            if failed_files:
                print("\nFailed files:")
                for file, error in failed_files:
                    print(f"  - {Path(file).name}: {error}")

            if successful_files:
                logger.info(f"Successfully processed {len(successful_files)} files")
                print(f"\n✓ Successfully processed {len(successful_files)} file(s)")

            if failed_files and not successful_files:
                logger.error("All files failed to process")
                print("\n✗ All files failed to process. Please check the errors above.")

            # Clear lists after processing
            for file_list in self.biglist:
                file_list.clear()

        proceed = str(input("\nConfirm Y to continue, N to go back: "))
        if proceed.upper() == "Y":
            process_translation_batch()
        else:
            self.edit_inputs(
                input_sheet,
                start_row,
                ilanguage,
                olanguage,
                input_column,
                filetype,
                glossfile,
                data_path,
                output_dir,
                api_keys
            )


class Begin:
    """
    Main entry point class for the translation module.
    """

    def modrun(
        self,
        data_path: str,
        ilanguage: str,
        olanguage: str,
        filetype: str,
        glossfile: str,
        sheet: str,
        input_column: List[str],
        start_row: int,
        output_dir: Path,
        api_keys: List[str],
    ) -> None:
        """
        Run the complete translation workflow.

        Args:
            data_path: Path to data directory or file
            ilanguage: Input language code
            olanguage: Output language code
            filetype: Type of files (Archive, Monument, Excel, etc.)
            glossfile: Path to glossary file
            sheet: Sheet name to translate
            input_column: List of column letters to translate
            start_row: Starting row number
            output_dir: Output directory for CSV files
            api_keys: List of DeepL API keys
        """
        load = LoadFolders()
        load.openfolder(data_path, filetype)
        load.runtransl(
            ilanguage,
            olanguage,
            data_path,
            filetype,
            glossfile,
            sheet,
            input_column,
            start_row,
            output_dir,
            api_keys,
        )
