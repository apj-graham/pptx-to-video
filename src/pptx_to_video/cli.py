import argparse
import os
from typing import List

import Levenshtein
from languages import Languages
from logger import logger


class ValidatePPTXFile(argparse.Action):
    """Custom argparse action to validate PPTX file path."""

    def __call__(
        self,
        parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: str,
        option_string: str | None = None,
    ) -> None:
        pptx_path: str = values
        if not os.path.isfile(pptx_path):
            parser.error(f"PPTX file '{pptx_path}' does not exist or is not a file.")
        elif not pptx_path.lower().endswith(".pptx"):
            parser.error(f"File '{pptx_path}' is not a valid PPTX file.")

        setattr(namespace, self.dest, pptx_path)


class ValidateLanguage(argparse.Action):
    """Custom argparse action to validate language codes."""

    def __call__(
        self,
        parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: list[str],
        option_string: str | None = None,
    ) -> None:

        def is_similar_levenshtein(s1, s2, threshold=1) -> bool:
            distance: int = Levenshtein.distance(s1, s2)
            return distance <= threshold

        languages_manager: Languages = Languages()
        available_language_names: List[str] = (
            languages_manager.get_available_language_names()
        )

        invalid_languages: List[str] = []
        for lang in values:
            if not languages_manager.is_language_supported(lang):
                for suggestion in available_language_names:
                    if is_similar_levenshtein(lang, suggestion):
                        raise ValueError(
                            f"Language code '{lang}' not found. Did you mean '{suggestion}'?"
                        )
                    else:
                        invalid_languages.append(lang)

        if invalid_languages:
            parser.error(
                f"Invalid language codes: {', '.join(invalid_languages)}. "
                f"Available languages are: {', '.join(available_language_names)}"
            )

        setattr(namespace, self.dest, values)


class ValidateOutputDir(argparse.Action):
    """Custom argparse action to validate output directory."""

    def __call__(
        self,
        parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: str,
        option_string: str | None = None,
    ) -> None:
        output_dir: str = values
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                logger.info(f"Created output directory: {output_dir}")
            except OSError as e:
                parser.error(f"Failed to create output directory '{output_dir}': {e}")
        elif not os.path.isdir(output_dir):
            parser.error(f"Output path '{output_dir}' is not a directory.")

        setattr(namespace, self.dest, output_dir)


def arg_parser() -> argparse.ArgumentParser:
    """Argument parser for the PPTX to video script.

    :return: Parsed arguments
    :rtype: argparse.Namespace
    """
    parser = argparse.ArgumentParser(
        description="Convert PPTX to video with narration."
    )

    parser.add_argument(
        "pptx_path",
        type=str,
        help="Path to the input PPTX file.",
        action=ValidatePPTXFile,
    )

    parser.add_argument(
        "--languages",
        type=str,
        nargs="+",
        default=["en"],
        help="List of language codes for translation (default: ['en']).",
        action=ValidateLanguage,
    )

    parser.add_argument(
        "--output_dir",
        type=str,
        default=os.path.join(os.getcwd(), "output_videos"),
        help="Directory to save the output videos (default: 'output_videos').",
        action=ValidateOutputDir,
    )

    return parser
