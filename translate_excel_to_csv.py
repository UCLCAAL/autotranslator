"""
Created on 1 Dec 2021

@author: turch
"""
from deep_translator import GoogleTranslator
import pandas as pd
import fnmatch as fn
from pathlib import Path
import traceback
import logging
import time
from typing import List, Dict, Optional, Tuple, Any
from config import (
    COMBINE_COLUMNS,
    CSV_DELIMITER,
    CSV_ENCODING,
    COLUMN_NAME_MAPPINGS
)
from utils import ensure_output_directory

logger = logging.getLogger(__name__)


class TranslationCache:
    """Cache for translations to avoid redundant API calls."""
    
    def __init__(self):
        """Initialize translation cache."""
        self.cache: Dict[Tuple[str, str, str], str] = {}  # (text, source_lang, target_lang) -> translation
    
    def get(self, text: str, source_lang: str, target_lang: str) -> Optional[str]:
        """Get cached translation if available."""
        key = (text, source_lang, target_lang)
        return self.cache.get(key)
    
    def set(self, text: str, source_lang: str, target_lang: str, translation: str) -> None:
        """Store translation in cache."""
        key = (text, source_lang, target_lang)
        self.cache[key] = translation
    
    def clear(self) -> None:
        """Clear the cache."""
        self.cache.clear()


class TranslMethods:
    """
    Methods for translating Excel data to CSV format.
    
    Handles data loading, translation, and CSV export with glossary support.
    """
    
    def __init__(self):
        """Initialize TranslMethods with instance-level lists to avoid shared state."""
        self.checked_untrans: List[Dict[str, Any]] = []
        self.translated: List[Dict[str, Any]] = []
        self.combineddata: List[Dict[str, Any]] = []
        self.untranslated: List[Dict[str, Any]] = []
        self.list_combined: List[Dict[str, Any]] = []
        self.rows = 0
        self.linked_cols: Dict[str, str] = {}
        self.transl_columns: List[str] = []
        self.russgloss: Optional[Dict[str, Any]] = None
        self.translator_instance: Optional[GoogleTranslator] = None
        self.translation_cache = TranslationCache()
        
    def _get_translator(self, source_lang: str, target_lang: str) -> GoogleTranslator:
        """
        Get or create translator instance for reuse.
        
        Args:
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            GoogleTranslator instance
        """
        if self.translator_instance is None:
            self.translator_instance = GoogleTranslator(source=source_lang, target=target_lang)
            logger.debug(f"Created translator instance: {source_lang} -> {target_lang}")
        return self.translator_instance
        
    def _translate_with_retry(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        max_retries: int = 3,
        retry_delay: float = 1.0
    ) -> str:
        """
        Translate text with retry logic for handling API rate limits and timeouts.
        
        Args:
            text: Text to translate
            source_lang: Source language code
            target_lang: Target language code
            max_retries: Maximum number of retry attempts
            retry_delay: Delay between retries in seconds
            
        Returns:
            Translated text
            
        Raises:
            Exception: If translation fails after all retries
        """
        # Check cache first
        cached = self.translation_cache.get(text, source_lang, target_lang)
        if cached:
            logger.debug(f"Cache hit for text: {text[:50]}...")
            return cached
        
        # Skip translation for empty strings
        if not text or str(text).strip() == '':
            return str(text)
        
        translator = self._get_translator(source_lang, target_lang)
        
        for attempt in range(max_retries):
            try:
                translation = translator.translate(str(text))
                # Cache the result
                self.translation_cache.set(text, source_lang, target_lang, translation)
                return translation
            except Exception as e:
                if attempt < max_retries - 1:
                    logger.warning(
                        f"Translation attempt {attempt + 1} failed: {str(e)}. "
                        f"Retrying in {retry_delay} seconds..."
                    )
                    time.sleep(retry_delay * (attempt + 1))  # Exponential backoff
                else:
                    error_msg = f"Translation failed after {max_retries} attempts: {str(e)}"
                    logger.error(error_msg)
                    raise Exception(error_msg) from e
        
        # Should never reach here, but just in case
        raise Exception("Translation failed unexpectedly")
        
    def load_data(
        self,
        filename: str,
        input_sheet: str,
        input_column: List[str],
        start_row: int,
        max_row: int
    ) -> None:
        """
        Load data from Excel file.
        
        Args:
            filename: Path to Excel file
            input_sheet: Name of sheet to read
            input_column: List of column letters to read
            start_row: Row number to start reading from (1-indexed)
            max_row: Maximum row number
        """
        logger.info(f"Loading data from {filename}, sheet: {input_sheet}")
        
        col = ','.join(input_column)  # joins column names as a string with , - needed to set the columns for the dataframe
        rows = int(max_row - start_row)  # number of rows is max_row minus start row
        self.rows = rows
        
        logger.info(f"Data length: {rows} rows. Starting row: {start_row}")
        
        skipline = int(start_row) - 1  # pandas uses 0-index, so skip lines before start_row
        try:
            data = pd.read_excel(
                filename,
                input_sheet,
                header=None,
                names=input_column,
                usecols=col,
                skiprows=skipline,
                nrows=rows
            )
            logger.debug(f"Data read from Excel:\n{data.head()}")
            data = data.fillna('')  # fill empty cells with empty strings
            
            # Get column names from the spreadsheet
            names = pd.read_excel(filename, input_sheet, header=1, usecols=col, nrows=1)
            headers = list(names.columns)
            transl_columns = []
            
            # Translate column headers
            translator = GoogleTranslator(source='ru', target='en')
            for h in headers:
                try:
                    transl_h = translator.translate(str(h))
                    # Map common column names
                    transl_h_lower = str(transl_h).lower()
                    for key, value in COLUMN_NAME_MAPPINGS.items():
                        if key in transl_h_lower:
                            transl_h = value
                            break
                    transl_columns.append(transl_h)
                except Exception as e:
                    logger.warning(f"Could not translate column header '{h}': {str(e)}. Using original.")
                    transl_columns.append(str(h))
            
            self.transl_columns = transl_columns
            linked_cols = dict(zip(input_column, transl_columns))
            self.linked_cols = linked_cols
            logger.debug(f"Column mappings: {linked_cols}")
            
            # Convert to list of dictionaries
            untrans_to_dict = data[input_column].to_dict(orient='records')
            untranslated = [
                {linked_cols.get(k, k): v for k, v in entry.items()}
                for entry in untrans_to_dict
            ]
            self.untranslated = untranslated
            
            logger.info("Data extracted successfully")
            
        except Exception as e:
            error_msg = f"Error loading data from {filename}: {str(e)}"
            logger.exception(error_msg)
            raise
    
    def create_glossary(self, ilang: str, glossfile: str) -> Dict[str, Any]:
        """
        Create glossary dictionary from Excel file.
        
        Args:
            ilang: Input language code
            glossfile: Path to glossary Excel file
            
        Returns:
            Dictionary mapping source terms to target terms
        """
        logger.info(f"Creating glossary from {glossfile} for language {ilang}")
        
        try:
            if ilang == 'ru':
                rus = pd.read_excel(glossfile, usecols='B:C')
                csv_path = Path("Rusgloss.csv")
                rus.to_csv(csv_path, index=None, header=True)
                rusdf = pd.read_csv(csv_path, encoding='utf-8')
                russgloss = rusdf.set_index('Russian').T.to_dict('list')
                self.russgloss = russgloss
                logger.info('Russian glossary created')
                return russgloss
            elif ilang == 'zh_CN':
                cn = pd.read_excel(glossfile, usecols='C:D')
                csv_path = Path('Chinese_gloss.csv')
                cn.to_csv(csv_path, index=None, header=True)
                cndf = pd.read_csv(csv_path, encoding='utf-8')
                cngloss = cndf.set_index('Chinese').T.to_dict('list')
                self.cngloss = cngloss
                logger.info('Chinese glossary created')
                return cngloss
            else:
                logger.warning(f"Unsupported language for glossary: {ilang}")
                return {'1': ['1']}
        except Exception as e:
            error_msg = f"Error creating glossary: {str(e)}"
            logger.exception(error_msg)
            raise
    
    def translator(self, ilanguage: str, olanguage: str, filetype: str) -> None:
        """
        Translate untranslated data using Google Translate API.
        
        Args:
            ilanguage: Input language code
            olanguage: Output language code
            filetype: Type of file being processed
        """
        # Use untranslated directly for Archive files (bypasses data_check)
        if filetype == 'Archive':
            untranslated = self.untranslated
        else:
            untranslated = self.checked_untrans if self.checked_untrans else self.untranslated

        logger.info('Starting translation - this may take a while')
        print('Now translating - this can take a while, make a cup of tea :)')

        try:
            n = 0
            for data_dict in untranslated:
                transl_data = {}
                for col, item in data_dict.items():
                    if col.lower() == 'caal_id':
                        transl_data[col] = item  # Don't translate CAAL ID
                    else:
                        transl_text = str(item)
                        try:
                            transltxt = self._translate_with_retry(
                                transl_text,
                                ilanguage,
                                olanguage
                            )
                            transl_data[col] = transltxt
                        except Exception as e:
                            logger.error(f"Translation failed for column {col}: {str(e)}")
                            transl_data[col] = transl_text  # Use original on failure
                
                self.translated.append(transl_data)
                n += 1
                if n % 10 == 0:  # Log every 10 entries
                    logger.info(f"{n} entries translated")
                    print(f"{n} entries translated")
                    
        except Exception as e:
            error_msg = f"Error during translation: {str(e)}"
            logger.exception(error_msg)
            print(f"ERROR: {error_msg}")
            raise
    
    def combinedata(self) -> None:
        """
        Combine untranslated and translated data for output.
        
        Creates final data structure with combined original/translated text
        for specified columns.
        """
        logger.info("Combining translated and original data")
        list_combined = []

        for data_dict in self.translated:
            combined_data = {}
            id_value = data_dict.get('CAAL_ID')
            
            # Find corresponding untranslated entry
            untrans = next(
                (d for d in self.untranslated if d.get('CAAL_ID') == id_value),
                {}
            )
            
            for col, item in data_dict.items():
                if col in COMBINE_COLUMNS and col in untrans:
                    original = untrans[col]
                    combined = f"{original} / {item}"
                    combined_data[col] = combined
                else:
                    combined_data[col] = item
            
            # Create final data structure
            # Note: This structure is specific to the current use case
            # May need to be made configurable in the future
            title_parts = [
                combined_data.get('CAAL_ID', ''),
                combined_data.get('Name', ''),
                combined_data.get('Title', '')
            ]
            title = ' / '.join(str(part) for part in title_parts if part)
            
            final_data = {
                'Title': title,
                'Description': combined_data.get('Description', ''),
                'Type': f"{combined_data.get('Type of material', '')}. {combined_data.get('Quantity', '')}"
            }
            
            list_combined.append(final_data)
        
        self.list_combined = list_combined
        logger.info(f"Combined {len(list_combined)} entries")
    
    def dict_to_csv(self, filename: str, output_dir: Path) -> None:
        """
        Export combined data to CSV file.
        
        Args:
            filename: Source filename (used for output naming)
            output_dir: Directory to save CSV file
        """
        csv_file = Path(filename).stem + '_translated_data.csv'
        
        if not self.list_combined:
            logger.warning("No data to export. Running combinedata() first.")
            self.combinedata()
        
        df = pd.DataFrame(self.list_combined)
        logger.debug(f"DataFrame to export:\n{df.head()}")
        
        # Ensure output directory exists
        output_path = ensure_output_directory(output_dir, create_subdir=False)
        csv_path = output_path / csv_file
        
        try:
            df.to_csv(
                csv_path,
                index=False,
                encoding=CSV_ENCODING,
                sep=CSV_DELIMITER
            )
            logger.info(f"Successfully produced CSV: {csv_path}")
            print(f'Successfully produced CSV: {csv_path}')
        except Exception as e:
            error_msg = f"Error writing CSV file: {str(e)}"
            logger.exception(error_msg)
            raise
    
    def reset(self) -> None:
        """Reset all instance variables for next translation batch."""
        self.untranslated.clear()
        self.checked_untrans.clear()
        self.translated.clear()
        self.combineddata.clear()
        self.list_combined.clear()
        self.translation_cache.clear()
        self.translator_instance = None
        logger.debug("Reset translation state")


class TranslateRun:
    """
    Class to run the translation workflow.
    """
    
    def __init__(self):
        """Initialize TranslateRun with TranslMethods instance."""
        self.transl = TranslMethods()
    
    def execute_translation(
        self,
        filename: str,
        input_sheet: str,
        input_column: List[str],
        start_row: int,
        max_row: int,
        ilanguage: str,
        olanguage: str,
        filetype: str,
        glossfile: str,
        output_dir: Path
    ) -> None:
        """
        Execute the complete translation workflow for a single file.
        
        Args:
            filename: Path to Excel file
            input_sheet: Sheet name to translate
            input_column: List of column letters to translate
            start_row: Starting row number
            max_row: Maximum row number
            ilanguage: Input language code
            olanguage: Output language code
            filetype: Type of file (Archive, Monument, etc.)
            glossfile: Path to glossary file
            output_dir: Output directory for CSV files
        """
        self.transl.load_data(filename, input_sheet, input_column, start_row, max_row)
        # data_check is currently bypassed for Archive files
        # self.transl.data_check(ilanguage, olanguage, filetype, glossfile)
        self.transl.translator(ilanguage, olanguage, filetype)
        self.transl.combinedata()
        self.transl.dict_to_csv(filename, output_dir)
    
    def runtransl(
        self,
        filename: str,
        input_sheet: str,
        input_column: List[str],
        start_row: int,
        max_row: int,
        ilanguage: str,
        olanguage: str,
        filetype: str,
        glossfile: str,
        output_dir: Path
    ) -> None:
        """
        Run translation and reset state afterwards.
        
        Args:
            filename: Path to Excel file
            input_sheet: Sheet name to translate
            input_column: List of column letters to translate
            start_row: Starting row number
            max_row: Maximum row number
            ilanguage: Input language code
            olanguage: Output language code
            filetype: Type of file
            glossfile: Path to glossary file
            output_dir: Output directory for CSV files
        """
        logger.info(f"Translating: {filename}\nColumns: {input_column}")
        print(f'Now translating: {filename}\nColumns: {input_column}')
        print('Note: If file contains forms/macros you will need to manually close/log into these when they pop-up during entry')
        
        self.execute_translation(
            filename,
            input_sheet,
            input_column,
            start_row,
            max_row,
            ilanguage,
            olanguage,
            filetype,
            glossfile,
            output_dir
        )
        self.transl.reset()
        
