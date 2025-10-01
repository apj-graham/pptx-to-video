import argparse
import os


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
    )

    parser.add_argument(
        "--languages",
        type=str,
        nargs="+",
        default=["en"],
        help="List of language codes for translation (default: ['en']).",
    )

    parser.add_argument(
        "--output_dir",
        type=str,
        default=os.path.join(os.getcwd(), "output_videos"),
        help="Directory to save the output videos (default: 'output_videos').",
    )

    return parser
