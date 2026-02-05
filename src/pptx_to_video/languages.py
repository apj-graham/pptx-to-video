import csv
from dataclasses import dataclass
from pathlib import Path
from pprint import pprint
from typing import List

import argostranslate.package

from pptx_to_video.logger import logger
from pptx_to_video.tts import PIPER_LANGUAGES


@dataclass
class LanguageInfo:
    """Represents information about a language including its name, ISO code, and Piper locale code."""

    name: str
    iso_code: str
    piper_locale_code: str


class Languages:
    """Manages language information and mappings for translation and text-to-speech functionality.

    This class handles loading ISO language codes, finding common languages supported by
    both Argos Translate and Piper TTS, and providing language information for the application.
    """

    ISO_PATH: Path = (
        Path(__file__).resolve().parent.parent.parent / "data" / "iso-codes.csv"
    )

    def __init__(self):
        logger.info("Initializing Languages manager...")
        self.code_to_language: dict[str, str] = self.load_iso_name_mapping()
        self.available_languages: List[LanguageInfo] = self._load_available_languages()
        logger.info(
            f"Loaded {len(self.available_languages)} available languages "
            "supported by both translation and TTS."
        )

    def get_available_language_names(self) -> List[str]:
        """Get a list of names for all available languages.

        Returns:
            List[str]: A list of language names that are supported by both
                      translation and text-to-speech systems.
        """
        return [lang.name.lower() for lang in self.available_languages]

    def load_iso_name_mapping(self) -> dict[str, str]:
        """Load ISO 639-1 language codes and their corresponding names from CSV file.

        Reads the iso-codes.csv file and creates a mapping from ISO codes to language names.

        Returns:
            dict[str, str]: A dictionary mapping ISO 639-1 codes to language names.
        """
        mapping: dict[str, str] = {}
        logger.info(f"Loading ISO language codes from {self.ISO_PATH}...")
        with open(self.ISO_PATH, newline="", encoding="utf-8") as csvfile:
            reader: csv.DictReader = csv.DictReader(csvfile)
            for row in reader:
                mapping[row["639-1"]] = row["name"].split(" (")[0]
        logger.info(f"Loaded {len(mapping)} ISO language codes.")
        return mapping

    def _get_argos_language_codes(self) -> set[str]:
        """Get language codes supported by Argos Translate for translation from English.

        Updates the package index and retrieves available translation packages
        that can translate from English to other languages.

        Returns:
            set[str]: A set of ISO language codes supported by Argos Translate.
        """
        logger.info("Fetching Argos Translate packages...")
        argostranslate.package.update_package_index()
        logger.info("Argos package list updated.")

        available_packages = argostranslate.package.get_available_packages()
        return set([pkg.to_code for pkg in available_packages if pkg.from_code == "en"])

    def _get_piper_language_codes(self) -> set[str]:
        """Get language codes supported by Piper TTS.

        Retrieves available Piper TTS languages and extracts the ISO 639-1 codes
        (first 2 characters) from the locale codes.

        Returns:
            set[str]: A set of ISO 639-1 language codes supported by Piper TTS.
        """

        return set([lang[:2] for lang in PIPER_LANGUAGES])

    def _iso_to_piper_locale(self, iso_code: str) -> str:
        """Convert an ISO 639-1 language code to a Piper locale code.

        Finds the first Piper locale that starts with the given ISO code.

        Args:
            iso_code (str): ISO 639-1 language code (e.g., 'en', 'es').

        Returns:
            str: The corresponding Piper locale code, or empty string if not found.
        """

        for locale in PIPER_LANGUAGES:
            if locale.startswith(iso_code):
                return locale
        return ""

    def _load_available_languages(self) -> List[LanguageInfo]:
        """Load information for languages supported by both Argos Translate and Piper TTS.

        Finds the intersection of languages supported by both systems and creates
        LanguageInfo objects with proper names and locale codes.

        Returns:
            List[LanguageInfo]: A list of LanguageInfo objects for supported languages.
        """
        argos_codes: set[str] = self._get_argos_language_codes()
        piper_codes: set[str] = self._get_piper_language_codes()

        common_codes: set[str] = argos_codes.intersection(piper_codes)

        languages: List[LanguageInfo] = []
        for code in common_codes:
            piper_locale: str = self._iso_to_piper_locale(code)
            languages.append(
                LanguageInfo(
                    name=self.code_to_language.get(code, "Unknown"),
                    iso_code=code,
                    piper_locale_code=piper_locale,
                )
            )
        return languages


if __name__ == "__main__":
    languages = Languages()
    pprint(languages.get_available_language_names())
