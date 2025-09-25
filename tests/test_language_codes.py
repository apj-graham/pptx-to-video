from unittest import mock

import pytest

from src.pptx_to_video.language_codes import language_to_code


class TestLanguageCodes:
    @pytest.fixture(autouse=True)
    def mock_translator(self):
        with mock.patch("deep_translator.GoogleTranslator") as MockTranslator:
            instance = MockTranslator.return_value
            instance.get_supported_languages.return_value = {
                "english": "en",
                "spanish": "es",
                "french": "fr",
            }
            yield

    def test_language_to_code_supported_language(self):
        """Test language_to_code with supported languages returns correct codes"""
        assert language_to_code("English") == "en"
        assert language_to_code("SPANISH") == "es"
        assert language_to_code("french") == "fr"

    def test_language_to_code_unsupported_language(self):
        """Test language_to_code with unsupported language raises ValueError"""
        with pytest.raises(ValueError) as excinfo:
            language_to_code("Klingon")
        assert "Language 'Klingon' not supported." in str(excinfo.value)
