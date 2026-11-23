"""
Created on 1 Dec 2021

@author: turch
"""
from deep_translator import DeeplTranslator
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
    COLUMN_NAME_MAPPINGS,
    LOCATION_COLUMN_ARCHIVE,
    CORE_DRIVE
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
    
    def __init__(self, api_keys: Optional[List[str]] = None):
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
        self.api_keys: List[str] = api_keys or []
        self.current_key_index: int = 0
        self.translator_instance: Optional[DeeplTranslator] = None
        self._active_lang_pair: Optional[Tuple[str, str]] = None
        self.translation_cache = TranslationCache()
        self.filetype: Optional[str] = None
        self.location_field: str = '__archive_location_raw__'

    def set_api_keys(self, api_keys: List[str]) -> None:
        """Set or update the list of DeepL API keys."""
        if not api_keys:
            raise ValueError("api_keys list cannot be empty.")
        if api_keys != self.api_keys:
            self.api_keys = api_keys
            self.current_key_index = 0
            self.translator_instance = None
        
    def _get_translator(self, source_lang: str, target_lang: str) -> DeeplTranslator:
        """
        Get or create DeepL translator instance for the current active API key.
        
        Args:
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            DeeplTranslator instance
        """
        if not self.api_keys:
            raise ValueError("No DeepL API keys provided. Pass api_keys before translating.")

        if self.current_key_index >= len(self.api_keys):
            raise RuntimeError("All DeepL API keys have been exhausted or failed.")

        if self.translator_instance is None or self._active_lang_pair != (source_lang, target_lang):
            current_key = self.api_keys[self.current_key_index]
            self.translator_instance = DeeplTranslator(
                api_key=current_key,
                source=source_lang,
                target=target_lang,
                use_free_api=current_key.endswith(":fx")
            )
            self._active_lang_pair = (source_lang, target_lang)
            logger.debug(
                f"Created DeepL translator instance ({source_lang} -> {target_lang}) "
                f"using key #{self.current_key_index + 1}"
            )
        return self.translator_instance
        
    def _translate_with_retry(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        retry_delay: float = 1.0
    ) -> str:
        """
        Translate text with automatic API key rotation on failure/quota exhaustion.
        
        Args:
            text: Text to translate
            source_lang: Source language code
            target_lang: Target language code
            retry_delay: Brief pause before switching keys
            
        Returns:
            Translated text
            
        Raises:
            Exception: If all API keys fail
        """
        # Check cache first
        cached = self.translation_cache.get(text, source_lang, target_lang)
        if cached:
            logger.debug(f"Cache hit for text: {text[:50]}...")
            return cached
        
        # Skip translation for empty strings
        if not text or str(text).strip() == '':
            return str(text)

        while self.current_key_index < len(self.api_keys):
            translator = self._get_translator(source_lang, target_lang)
            try:
                translation = translator.translate(str(text))
                self.translation_cache.set(text, source_lang, target_lang, translation)
                return translation
            except Exception as e:
                logger.warning(
                    f"DeepL key #{self.current_key_index + 1} failed: {str(e)}. "
                    f"Rotating to next key..."
                )
                print(f"Key #{self.current_key_index + 1} failed ({e}). Switching to next key...")
                self.current_key_index += 1
                self.translator_instance = None
                if self.current_key_index < len(self.api_keys):
                    time.sleep(retry_delay)

        error_msg = "Translation failed: all DeepL API keys exhausted or failed."
        logger.error(error_msg)
        raise RuntimeError(error_msg)
        
    def load_data(
        self,
        filename: str,
        input_sheet: str,
        input_column: List[str],
        start_row: int,
        max_row: int,
        filetype: str,
        ilanguage: str = 'ru',
        olanguage: str = 'en'
    ) -> None:
        """
        Load data from Excel file.
        
        Args:
            filename: Path to Excel file
            input_sheet: Name of sheet to read
            input_column: List of column letters to read
            start_row: Row number to start reading from (1-indexed)
            max_row: Maximum row number
            filetype: Type of file being processed
            ilanguage: Input language code for header translation
            olanguage: Output language code for header translation
        """
        logger.info(f"Loading data from {filename}, sheet: {input_sheet}")
        
        all_columns = list(input_column)
        if filetype == 'Archive' and LOCATION_COLUMN_ARCHIVE not in all_columns:
            all_columns.append(LOCATION_COLUMN_ARCHIVE)

        col = ','.join(all_columns)
        rows = int(max_row - start_row)
        self.rows = rows
        
        logger.info(f"Data length: {rows} rows. Starting row: {start_row}")
        
        skipline = int(start_row) - 1
        try:
            data = pd.read_excel(
                filename,
                input_sheet,
                header=None,
                names=all_columns,
                usecols=col,
                skiprows=skipline,
                nrows=rows
            )
            logger.debug(f"Data read from Excel:\n{data.head()}")
            data = data.fillna('')
            
            # Get column names from the spreadsheet
            names = pd.read_excel(filename, input_sheet, header=1, usecols=col, nrows=1)
            headers = list(names.columns)
            transl_columns = []
            
            # Translate column headers using the rotating DeepL keys
            for h in headers:
                try:
                    transl_h = self._translate_with_retry(str(h), ilanguage, olanguage)
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
            untrans_to_dict = data[all_columns].to_dict(orient='records')
            untranslated = []
            for entry in untrans_to_dict:
                converted_entry = {linked_cols.get(k, k): v for k, v in entry.items()}
                if filetype == 'Archive':
                    converted_entry[self.location_field] = entry.get(LOCATION_COLUMN_ARCHIVE, '')
                untranslated.append(converted_entry)
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
    
    def translator(
        self,
        ilanguage: str,
        olanguage: str,
        filetype: str,
        api_keys: Optional[List[str]] = None
    ) -> None:
        """
        Translate untranslated data using DeepL API with automatic key rotation.
        
        Args:
            ilanguage: Input language code
            olanguage: Output language code
            filetype: Type of file being processed
            api_keys: Optional list of DeepL API keys (if not already set on instance)
        """
        if api_keys:
            self.set_api_keys(api_keys)

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
                    if col.lower() == 'caal_id' or col == self.location_field:
                        transl_data[col] = item
                    else:
                        transl_text = str(item)
                        try:
                            transltxt = self._translate_with_retry(
                                transl_text,
                                ilanguage,
                                olanguage
                            )
                            transl_data[col] = transltxt
                        except RuntimeError:
                            # Re-raise immediately if all API keys are exhausted
                            raise
                        except Exception as e:
                            logger.error(f"Translation failed for column {col}: {str(e)}")
                            transl_data[col] = transl_text
                
                self.translated.append(transl_data)
                n += 1
                if n % 10 == 0:
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
        """
        logger.info("Combining translated and original data")
        list_combined = []

        for data_dict in self.translated:
            combined_data = {}
            id_value = data_dict.get('CAAL_ID')
            
            untrans = next(
                (d for d in self.untranslated if d.get('CAAL_ID') == id_value),
                {}
            )
            
            for col, item in data_dict.items():
                if col in COMBINE_COLUMNS and col in untrans:
                    original = untrans[col]
                    combined = f"{original} - {item}"
                    combined_data[col] = combined
                else:
                    combined_data[col] = item
            
            title_parts = [
                combined_data.get('CAAL_ID', ''),
                combined_data.get('Name', ''),
                combined_data.get('Title', '')
            ]
            title = ' - '.join(str(part) for part in title_parts if part)
            
            final_data = {
                'Title': title,
                'Description': combined_data.get('Description', ''),
                'Type': f"{combined_data.get('Type of material', '')}. {combined_data.get('Quantity', '')}"
            }
            if self.filetype == 'Archive':
                raw_location = combined_data.get(self.location_field, '')
                if raw_location:
                    location_suffix = str(raw_location).strip().strip('/')
                    if location_suffix:
                        if CORE_DRIVE.endswith('/'):
                            location_full = f"{CORE_DRIVE.rstrip('/')}/{location_suffix}"
                        else:
                            location_full = f"{CORE_DRIVE}/{location_suffix}"
                        final_data['Location'] = f"Location: {location_full}"
            
            list_combined.append(final_data)
        
        self.list_combined = list_combined
        logger.info(f"Combined {len(list_combined)} entries")
    
    def dict_to_csv(self, filename: str, output_dir: Path) -> None:
        """
        Export combined data to CSV file.
        """
        csv_file = Path(filename).stem + '_translated_data.csv'
        
        if not self.list_combined:
            logger.warning("No data to export. Running combinedata() first.")
            self.combinedata()
        
        df = pd.DataFrame(self.list_combined)
        logger.debug(f"DataFrame to export:\n{df.head()}")
        
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
        """Reset per-file state while keeping active API key progress."""
        self.untranslated.clear()
        self.checked_untrans.clear()
        self.translated.clear()
        self.combineddata.clear()
        self.list_combined.clear()
        self.translation_cache.clear()
        self.translator_instance = None
        self._active_lang_pair = None
        logger.debug("Reset translation state")


class TranslateRun:
    """
    Class to run the translation workflow.
    """
    
    def __init__(self, api_keys: Optional[List[str]] = None):
        """Initialize TranslateRun with TranslMethods instance."""
        self.transl = TranslMethods(api_keys=api_keys)
    
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
        output_dir: Path,
        api_keys: Optional[List[str]] = None
    ) -> None:
        """
        Execute the complete translation workflow for a single file.
        """
        if api_keys:
            self.transl.set_api_keys(api_keys)

        self.transl.filetype = filetype
        self.transl.load_data(
            filename,
            input_sheet,
            input_column,
            start_row,
            max_row,
            filetype,
            ilanguage=ilanguage,
            olanguage=olanguage
        )
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
        output_dir: Path,
        api_keys: Optional[List[str]] = None
    ) -> None:
        """
        Run translation and reset state afterwards.
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
            output_dir,
            api_keys=api_keys
        )
        self.transl.reset()