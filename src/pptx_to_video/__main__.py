from pptx_to_video.cli import arg_parser
from pptx_to_video.logger import logger
from pptx_to_video.powerpoint import PowerPoint
from pptx_to_video.video_constructer import VideoConstructer


def main() -> None:
    """Main function to run the PPTX to video conversion script"""
    parser = arg_parser()
    args = parser.parse_args()

    pptx_path: str = args.pptx_path
    output_dir: str = args.output_dir
    languages: list[str] = args.languages

    logger.info(f"Input PPTX path: {pptx_path}")
    logger.info(f"Output directory: {output_dir}")
    logger.info(f"Languages for translation: {languages}")

    powerpoint = PowerPoint(pptx_path, output_dir)

    for lang in languages:
        logger.info(f"Processing language: {lang}")
        video_constructer = VideoConstructer(powerpoint, output_dir, lang)

        translated_texts = video_constructer.translate_notes()
        audio_paths = video_constructer.synthesize_speech(translated_texts)
        video_constructer.assemble_video(audio_paths)

    logger.info("All videos created successfully.")


if __name__ == "__main__":
    main()
