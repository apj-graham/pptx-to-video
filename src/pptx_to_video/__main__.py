from argparse import ArgumentParser, Namespace

from pptx_to_video.cli import arg_parser
from pptx_to_video.logger import logger

"""Inputs:
- PPTX file path
- Output directory
- List of target languages
"""


def main() -> None:
    """Main function to run the PPTX to video conversion script"""
    parser: ArgumentParser = arg_parser()
    args: Namespace = parser.parse_args()

    pptx_path: str = args.pptx_path
    output_dir: str = args.output_dir
    languages: list[str] = args.languages

    logger.info(f"Input PPTX path: {pptx_path}")
    logger.info(f"Output directory: {output_dir}")
    logger.info(f"Languages for translation: {languages}")


if __name__ == "__main__":
    main()
