import asyncio
from typing import List

import argostranslate.package
import argostranslate.translate
from tqdm import tqdm

from pptx_to_video.languages import LanguageInfo
from pptx_to_video.logger import logger


class TextTranslationEngine:
    """Class to manage text translations using Argos Translation"""

    def __init__(self, to_languages: List[LanguageInfo]):
        self.to_languages: List[LanguageInfo] = to_languages

        logger.info("Initializing TextTranslationEngine...")
        logger.info("Updating Argos Translate package index...")
        argostranslate.package.update_package_index()
        self.available_packages = argostranslate.package.get_available_packages()
        logger.info(
            f"Argos Translate package index updated with {len(self.available_packages)} packages."
        )

        logger.info("Installing required Argos Translate packages...")
        self._install_argos_packages()
        self.installed_languages = argostranslate.package.get_installed_packages()
        logger.info(
            f"Installed Argos Translate packages for languages: {', '.join([pkg.to_code for pkg in self.installed_languages])}"
        )

    def _install_argos_packages(self, max_workers: int = 3) -> None:
        """Download and install argostranslate translation packages for all supplied languages

        :param max_workers: number of packages to download at once, defaults to 3
        :type max_workers: int, optional
        :return: None
        :rtype: None
        """
        return asyncio.run(self._async_install_argos_packages(max_workers=max_workers))

    async def _async_install_argos_packages(self, max_workers: int = 3) -> None:
        """Asynchronously download and install argostranslate translation packages for all supplied languages

        :param max_workers: number of packages to download at once, defaults to 3
        :type max_workers: int, optional
        :return: None
        :rtype: None
        """

        packages_to_install: List[argostranslate.package.Package] = [
            self._get_translation_package(lang_info) for lang_info in self.to_languages
        ]

        # Only install packages that are not already installed
        packages_to_install: List[argostranslate.package.Package] = [
            pkg
            for pkg in packages_to_install
            if pkg not in argostranslate.package.get_installed_packages()
        ]

        if packages_to_install:
            # Create semaphore to limit concurrency
            semaphore = asyncio.Semaphore(max_workers)

            async def install_with_semaphore(package: argostranslate.package.Package):
                async with semaphore:
                    argostranslate.package.install_from_path(package.download())

            tasks = [install_with_semaphore(pkg) for pkg in packages_to_install]

            # Progress bar with proper updates
            with tqdm(
                total=len(tasks),
                desc="Installing Argos Translate packages",
                unit="package",
            ) as pbar:
                for coro in asyncio.as_completed(tasks):
                    try:
                        await coro
                        pbar.update(1)
                    except Exception as e:
                        print(f"Error installing package: {e}")
                        pbar.update(1)

    def _get_translation_package(
        self, to_language: LanguageInfo
    ) -> argostranslate.package.Package:
        """_summary_

        :param to_language: _description_
        :type to_language: LanguageInfo
        :raises ValueError: _description_
        :return: _description_
        :rtype: argostranslate.package.Package
        """

        try:
            package = list(
                filter(
                    lambda pkg: pkg.from_code == "en"
                    and pkg.to_code == to_language.iso_code,
                    self.available_packages,
                )
            )[0]
        except IndexError:
            raise ValueError(
                f"No translation package found for English -> {to_language.name}"
            )

        return package

    def translate_text(self, text: str, target_language: LanguageInfo) -> str:
        """Translate English text to target language

        :param text: Text to be translated
        :type text: str
        :param target_language: Language to translate text to
        :type target_language: LanguageInfo
        :return: Translated text
        :rtype: str
        """
        return argostranslate.translate.translate(
            text, from_code="en", to_code=target_language.iso_code
        )


if __name__ == "__main__":
    # Example usage
    languages = [LanguageInfo(name="Spanish", iso_code="es", piper_locale_code="es-ES")]
    translator = TextTranslationEngine(to_languages=languages)
    translated = translator.translate_text(
        "Hello, world!", target_language=languages[0]
    )
    print(translated)
