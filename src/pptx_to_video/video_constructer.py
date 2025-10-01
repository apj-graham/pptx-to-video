import os
from typing import List, Union

from deep_translator import GoogleTranslator
from gtts import gTTS
from moviepy import AudioFileClip, ImageClip, concatenate_videoclips

from pptx_to_video.language_codes import language_to_code
from pptx_to_video.logger import logger
from pptx_to_video.powerpoint import PowerPoint


class VideoConstructer:
    """Class to construct video from PowerPoint slides and notes"""

    def __init__(
        self,
        powerpoint: PowerPoint,
        output_dir: Union[str, os.PathLike],
        language: str,
    ) -> None:
        self.powerpoint: PowerPoint = powerpoint
        self.output_dir: Union[str, os.PathLike] = output_dir
        self.language: str = language  # e.g., 'English', 'Spanish', 'French'
        self.language_code: str = language_to_code(
            self.language
        )  # e.g., 'en', 'es', 'fr'
        self.tmp_dir: Union[str, os.PathLike] = os.path.join(
            self.output_dir, "tmp_assets"
        )
        os.makedirs(self.tmp_dir, exist_ok=True)

    def translate_notes(self) -> List[str]:
        """Translate notes to the target language

        :return: Note texts translated to target language
        :rtype: List[str]
        """
        logger.info("Starting translation for all notes...")
        results = []
        for note in self.powerpoint.slide_notes:
            logger.info(
                f"Translating notes for Slide {note['index'] + 1} to {self.language_code}"
            )
            if not note["notes"].strip():
                results.append("")
            else:
                res = GoogleTranslator(
                    source="auto", target=self.language_code
                ).translate(note["notes"])
                results.append(res)
        logger.info("Finished translating all notes")
        return results

    def synthesize_speech(self, texts: List[str]) -> List[Union[str, os.PathLike]]:
        """Synthesize speech for each text using Google TTS

        :param texts: List of texts to synthesize
        :type texts: List[str]
        :return: List of paths to synthesized audio files
        :rtype: List[Path]
        """
        logger.info("Starting speech synthesis for all notes...")
        audio_paths = []
        for i, text in enumerate(texts):
            if not text.strip():
                logger.info(f"Skipping empty notes for Slide {i + 1}")
                audio_paths.append(None)
                continue

            out_path = os.path.join(
                self.tmp_dir, f"slide_{i + 1:03d}_narration_{self.language_code}.mp3"
            )

            logger.info(f"Synthesizing speech for Slide {i + 1} to {out_path}")
            tts = gTTS(text=text, lang=self.language_code)
            tts.save(out_path)
            audio_paths.append(out_path)

        logger.info("Finished speech synthesis for all notes")
        return audio_paths

    def assemble_video(
        self, audio_files: List[Union[str, os.PathLike]]
    ) -> Union[str, os.PathLike]:
        """Assemble video from slide images and corresponding audio files

        :param audio_files: List of paths to audio files (can be None for slides without narration)
        :type audio_files: List[Union[str, os.PathLike]]
        :return: Path to the final video file
        :rtype: Union[str, os.PathLike]
        """

        logger.info("Starting video assembly...")

        clips = []
        for i, (img_path, audio_path) in enumerate(
            zip(self.powerpoint.slide_image_paths, audio_files)
        ):
            logger.info(f"Creating Video clip for Slide {i + 1}")

            slide_clip = self._construct_image_clip(
                img_path=img_path,
                audio_path=audio_path,
            )
            clips.append(slide_clip)

        final_video = concatenate_videoclips(clips, method="compose")
        output_video_path = os.path.join(
            self.output_dir, f"presentation_video_{self.language_code}.mp4"
        )
        final_video.write_videofile(
            output_video_path,
            codec="libx264",
            audio_codec="libmp3lame",
            audio=True,
            fps=24,
        )

        logger.info(f"Video assembly complete. Video saved to {output_video_path}")

        logger.info("Cleaning up temporary files...")
        self._tmp_dir_cleanup()
        logger.info("Temporary files cleaned up.")

        return output_video_path

    def _construct_image_clip(
        self,
        img_path: Union[str, os.PathLike],
        audio_path: Union[str, os.PathLike, None],
    ) -> ImageClip:
        """Helper method to create an ImageClip with specified duration

        :param img_path: Path to the image file
        :type img_path: Union[str, os.PathLike]
        :param duration: Duration for which the image should be displayed
        :type duration: float
        :return: ImageClip object
        :rtype: ImageClip
        """
        if audio_path and os.path.isfile(audio_path):
            audio_clip = AudioFileClip(audio_path)
            slide_clip = ImageClip(img_path, duration=audio_clip.duration).with_audio(
                audio_clip
            )
        else:
            logger.info(f"No audio for image {img_path}, using default duration")
            slide_clip = ImageClip(img_path, duration=5.0)  # default 5s if no audio
        return slide_clip

    def _tmp_dir_cleanup(self) -> None:
        """Helper method to clean up temporary directory"""
        for root, dirs, files in os.walk(self.tmp_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(self.tmp_dir)
