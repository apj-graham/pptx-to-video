from dataclasses import dataclass
from pathlib import Path
from typing import List, Union

from pptx_to_video.languages import LanguageInfo, Languages
from pptx_to_video.logger import logger
from pptx_to_video.powerpoint import PowerPoint
from pptx_to_video.translate import Translator
from pptx_to_video.tts import TextToSpeechEngine
from pptx_to_video.video_constructer import VideoConstructer

# 1. Extract slide images and notes from PPTX
# 2. For each language:
#   a. Translate notes
#   b. Generate audio from translated notes -
#   c. Construct video with slide images, translated notes as subtitles, and generated audio


@dataclass
class SlideContent:
    image_path: Union[str, Path]
    note_text: str
    translated_note_text: str
    audio_path: Union[str, Path]
    slide_number: int
    language: LanguageInfo


# Create SlideContent intances for each combination of slide and language
# Populate instances with translations
# Populate with audio paths after generating audio
# Then pass list of SlideContent instances to VideoConstructer to create videos (filter on language)


def translate_note(note: str, target_language: str) -> str:

    # translate note to target language
    # write to file
    # return file path
    pass


def create_note_audio(note: str, language: LanguageInfo) -> str:
    # read note path
    # generate audio from note text using TTS engine
    # write to file
    # return file path
    pass


class OrderedMethodRunner:
    _methods = []

    @classmethod
    def register(cls, func):
        cls._methods.append(func)
        return func

    def run(self):
        for method in self._methods:
            method(self)


class Orchestrater(OrderedMethodRunner):

    def __init__(self, pptx_path: str, output_dir: str, languages: list[LanguageInfo]):
        self.output_dir: str = output_dir
        self.languages: list[LanguageInfo] = languages

        # Extracts slides and images from PPTX
        self.pptx_engine = PowerPoint(pptx_path, output_dir=output_dir)

    @OrderedMethodRunner.register
    def first(self):
        print("First")

    @OrderedMethodRunner.register
    def second(self):
        print("Second")
