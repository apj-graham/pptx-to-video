import os
import platform
from typing import Union

from pptx_to_video.powerpoint_engine import PowerPointEngine


class PowerPoint:
    def __init__(
        self, pptx_path: Union[str, os.PathLike], output_dir: Union[str, os.PathLike]
    ) -> None:
        self.pptx_path: Union[str, os.PathLike] = pptx_path
        self.output_dir: Union[str, os.PathLike] = output_dir

        if platform.system() == "Windows":
            from pptx_to_video.powerpoint_engine import WindowsPowerPointEngine

            self.engine: PowerPointEngine = WindowsPowerPointEngine(self.pptx_path)

        elif platform.system() == "Darwin" or platform.system() == "Linux":
            from pptx_to_video.powerpoint_engine import LibreOfficePowerPointEngine

            self.engine: PowerPointEngine = LibreOfficePowerPointEngine(self.pptx_path)

        else:
            raise EnvironmentError(
                "Unsupported operating system for PowerPoint processing."
            )

        self.slide_image_paths = self.engine.export_slides_as_images(self.output_dir)
        self.slide_notes = self.engine.extract_notes_from_pptx()
