import os
import platform
from unittest import mock

import pytest

from src.powerpoint import PowerPoint
from src.powerpoint_engine import LibreOfficePowerPointEngine, WindowsPowerPointEngine

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUTS_PATH = os.path.join(CURRENT_DIR, "fixtures")
OUTPUTS_PATH = os.path.join(CURRENT_DIR, "outputs")


class TestPowerPoint:
    """Tests for the PowerPoint class"""

    @mock.patch("platform.system", return_value="UnsupportedOS")
    def test_init_invalid_platform(self, mock_platform):
        """Test initialization on an unsupported platform"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        output_dir = OUTPUTS_PATH

        with pytest.raises(EnvironmentError):
            PowerPoint(pptx_path, output_dir)

    @pytest.mark.skipif(
        platform.system() == "Windows",
        reason="LibreOfficePowerPointEngine tests can only run on Linux or MacOS",
    )
    @pytest.mark.parametrize(
        "platform_name, expected",
        [
            ("Darwin", LibreOfficePowerPointEngine),
            ("Linux", LibreOfficePowerPointEngine),
        ],
    )
    @mock.patch("platform.system")
    @mock.patch("shutil.which", return_value=True)
    def test_init_libre_platform(
        self, mock_which, mock_platform, platform_name, expected
    ):
        """Test initialization on supported platforms"""
        mock_platform.return_value = platform_name

        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        output_dir = OUTPUTS_PATH
        ppt = PowerPoint(pptx_path, output_dir)

        assert isinstance(ppt.engine, expected)
        assert ppt.pptx_path == pptx_path
        assert ppt.output_dir == output_dir

    @pytest.mark.skipif(
        platform.system() != "Windows",
        reason="WindowsPowerPointEngine tests can only run on Windows",
    )
    @mock.patch("win32com.client.Dispatch")
    def test_init_valid_platform(self, mock_dispatch):
        """Test initialization on supported platforms"""
        mock_dispatch.return_value = mock.MagicMock()
        mock_dispatch.return_value.Visible = False

        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        output_dir = OUTPUTS_PATH
        ppt = PowerPoint(pptx_path, output_dir)

        assert isinstance(ppt.engine, WindowsPowerPointEngine)
        assert ppt.pptx_path == pptx_path
        assert ppt.output_dir == output_dir
