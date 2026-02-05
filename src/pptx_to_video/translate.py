from pprint import pprint

import argostranslate
from argostranslate import package, translate

class TextTranslationEngine:
    def __init__(self):
        self.available_languages = self._load_available_languages()

    def get_supported_language_names(self) -> List[str]:
        """Get a list of supported language names.

         Retrieves the names of languages that are supported by both the translation
         and text-to-speech systems.

         Returns:
             List[str]: A list of language names that are supported by both
                       translation and text-to-speech systems.
         """
        return [lang.name for lang in self.available_languages]

    def _get_argos_language_codes(self) -> set[str]:
        """Get language codes supported by Argos Translate.

        Fetches the list of available translation packages from Argos Translate and
        extracts the target language codes (to_code) for packages that translate
        from English (from_code == "en").
