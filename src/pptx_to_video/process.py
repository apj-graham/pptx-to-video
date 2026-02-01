from pathlib import Path
from typing import List, Union

from pptx_to_video.powerpoint import PowerPoint
from pptx_to_video.logger import logger
from pptx_to_video.video_constructer import VideoConstructer
from pptx_to_video.language_codes import language_to_code


class Process:
    """Orchestrate production of subtitled and translated videos"""

    def __init__(
        self,
        pptx_path: Union[str, Path],
        output_dir: Union[str, Path],
        languages: list[str],
    ) -> None:
        self.pptx_path = pptx_path
        self.output_dir = output_dir
        self.languages = languages

        self.powerpoint = PowerPoint(
            pptx_path=self.pptx_path, output_dir=self.output_dir
        )

        self.slide_images: List[str] = self.powerpoint.slide_image_paths
        self.slide_notes: List[dict] = self.powerpoint.slide_notes

        self.language_codes: List[str] = self.languages