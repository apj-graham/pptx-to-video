import os
import platform
from unittest import mock

import pytest

from src.pptx_to_video.powerpoint_engine import (
    LibreOfficePowerPointEngine,
    PowerPointEngine,
    WindowsPowerPointEngine,
)

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUTS_PATH = os.path.join(CURRENT_DIR, "fixtures")
OUTPUTS_PATH = os.path.join(CURRENT_DIR, "outputs")


class TestPowerPointEngine:
    """Tests for PowerPointEngine class"""

    def test_cannot_instantiate_abstract_base_class(self):
        """Test that NotImplementedError is raised for base class method"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        with pytest.raises(TypeError):
            PowerPointEngine(pptx_path)


@pytest.mark.skipif(
    platform.system() != "Linux" or platform.system() != "Darwin",
    reason="LibreOfficePowerPointEngine tests can only run on Linux or MacOS",
)
class TestLibreOfficePowerPointEngine:
    """Tests for LibreOfficePowerPointEngine class"""

    @pytest.mark.parametrize("platform_name", ["Linux", "Darwin"])
    @mock.patch("platform.system")
    def test_init_correct_platform(self, mock_platform, platform_name):
        """Test initialization on non-Windows platform"""
        mock_platform.return_value = platform_name
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        engine = LibreOfficePowerPointEngine(pptx_path)
        assert isinstance(engine, LibreOfficePowerPointEngine)
        assert engine.pptx_path == pptx_path

    @mock.patch("platform.system", return_value="Windows")
    def test_init_incorrect_platform(self, mock_platform):
        """Test initialization on Windows platform raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        with pytest.raises(EnvironmentError):
            LibreOfficePowerPointEngine(pptx_path)

    def test_init_invalid_file_extension(self):
        """Test initialization with invalid file extension raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "invalid.txt")
        with pytest.raises(ValueError):
            LibreOfficePowerPointEngine(pptx_path)

    def test_init_nonexistent_file(self):
        """Test initialization with nonexistent file raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "nonexistent.pptx")
        with pytest.raises(FileNotFoundError):
            LibreOfficePowerPointEngine(pptx_path)

    @mock.patch("shutil.which", return_value=None)
    def test_init_libreoffice_not_installed(self, mock_which):
        """Test initialization when LibreOffice is not installed raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        with pytest.raises(EnvironmentError):
            LibreOfficePowerPointEngine(pptx_path)

    def test_extract_notes_from_pptx(self):
        """Test extracting notes from a sample pptx file"""
        # Create a sample pptx file with notes
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        engine = LibreOfficePowerPointEngine(pptx_path)

        notes_data = engine.extract_notes_from_pptx()

        assert len(notes_data) == 3
        assert notes_data[0]["index"] == 0
        assert notes_data[0]["notes"] == "These are notes for slide 1."
        assert notes_data[1]["index"] == 1
        assert notes_data[1]["notes"] == ""
        assert notes_data[2]["index"] == 2
        assert notes_data[2]["notes"] == "Notes for slide 3."


@pytest.mark.skipif(
    platform.system() != "Windows",
    reason="WindowsPowerPointEngine tests can only run on Windows",
)
class TestWindowsPowerPointEngine:
    """Tests for WindowsPowerPointEngine class"""

    def test_init_correct_platform(self):
        """Test initialization on Windows platform"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        engine = WindowsPowerPointEngine(pptx_path)
        assert isinstance(engine, WindowsPowerPointEngine)
        assert engine.pptx_path == pptx_path

    @mock.patch("platform.system", return_value="Linux")
    def test_init_incorrect_platform(self, mock_platform):
        """Test initialization on non Windows platform raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        with pytest.raises(EnvironmentError):
            WindowsPowerPointEngine(pptx_path)

    def test_init_invalid_file_extension(self):
        """Test initialization with invalid file extension raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "invalid.txt")
        with pytest.raises(ValueError):
            WindowsPowerPointEngine(pptx_path)

    def test_init_nonexistent_file(self):
        """Test initialization with nonexistent file raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "nonexistent.pptx")
        with pytest.raises(FileNotFoundError):
            WindowsPowerPointEngine(pptx_path)

    @mock.patch("win32com.client.Dispatch", side_effect=Exception)
    def test_init_powerpoint_not_installed(self, mock_dispatch):
        """Test initialization when PowerPoint is not installed raises error"""
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        with pytest.raises(RuntimeError):
            WindowsPowerPointEngine(pptx_path)

    def test_extract_notes_from_pptx(self):
        """Test extracting notes from a sample pptx file"""
        # Create a sample pptx file with notes
        pptx_path = os.path.join(INPUTS_PATH, "test.pptx")
        engine = LibreOfficePowerPointEngine(pptx_path)

        notes_data = engine.extract_notes_from_pptx()

        assert len(notes_data) == 3
        assert notes_data[0]["index"] == 0
        assert notes_data[0]["notes"] == "These are notes for slide 1."
        assert notes_data[1]["index"] == 1
        assert notes_data[1]["notes"] == ""
        assert notes_data[2]["index"] == 2
        assert notes_data[2]["notes"] == "Notes for slide 3."
