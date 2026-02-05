import asyncio
import random
import sys
from contextlib import contextmanager
from io import StringIO
from pathlib import Path
from pprint import pprint
from typing import Any, List, Optional, TextIO
from wave import Wave_write

from piper.download_voices import download_voice, list_voices
from piper.voice import PiperVoice
from tqdm import tqdm

from pptx_to_video.logger import logger


@contextmanager
def stdout_capture() -> TextIO:
    """Capture stdout output within the context manager.

    :yield: Captured stdout as StringIO
    :rtype: _type_
    """
    old_stdout: TextIO | Any = sys.stdout
    sys.stdout: StringIO = StringIO()
    try:
        yield sys.stdout
    finally:
        sys.stdout: TextIO | Any = old_stdout


def get_piper_languages() -> List[str]:
    """Get all available languages provided by PiperTTS

    :return: _description_
    :rtype: _type_
    """
    logger.info("Fetching Piper TTS available languages...")
    with stdout_capture() as captured_output:
        list_voices()
    logger.info(f"Updated Piper TTS languages.")

    return captured_output.getvalue().splitlines()


PIPER_LANGUAGES: List[str] = get_piper_languages()


class TextToSpeechEngine:

    VOICES_DIR: Path = (
        Path(__file__).resolve().parent.parent.parent / "data" / "piper_voices"
    )

    def __init__(self) -> None:
        self.VOICES_DIR.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def list_available_voices(language: Optional[str] = None) -> List[str]:
        """
        List all available TTS voices, optionally filtered by language.

        :param self: Description
        :param language: Language ISO code to filter by (e.g., 'en', 'es', 'fr')
        :type language: Optional[str]
        :return: List of available voice names
        :rtype: List[str]
        """

        if language:
            return list(
                filter(lambda voice: language.lower() in voice.lower(), PIPER_LANGUAGES)
            )

        return PIPER_LANGUAGES

    def list_available_languages(self) -> set[str]:
        """
        List all available TTS languages.

        :param self: Description
        :return: List of available language ISO codes
        :rtype: List[str]
        """
        voices: List[str] = self.list_available_voices()
        return set([voice[:5] for voice in voices])

    def get_voice(self, language: str, pitch: Optional[str] = None) -> str:
        """Get PiperVoice by locale code and pitch

        :param language: Piper locale code (e.g., 'en_US', 'de_DE')
        :type language: str
        :param pitch: Desired pitch in voice name (e.g. "low", "medium", "high")
        :type pitch: Optional[str]
        :raises ValueError: No voices found for the specified language
        :return: PiperVoice name
        :rtype: PiperVoice
        """
        voices: List[str] = self.list_available_voices(language)

        if not voices:
            raise ValueError(f"No voices found for language '{language}'")

        if pitch:
            for voice in voices:
                if pitch.lower() in voice.lower():
                    return voice

        return random.choice(voices)

    async def download_piper_voice(self, voice_name: str) -> str:
        """
        Download a specific Piper voice.

        :param voice_name: Name of the voice to download
        :type voice_name: str
        """
        loop: asyncio.AbstractEventLoop = asyncio.get_event_loop()
        await loop.run_in_executor(
            None,
            lambda: download_voice(voice=voice_name, download_dir=self.VOICES_DIR),
        )
        return voice_name

    async def _aysync_download_piper_voices(self, max_workers: int = 3) -> List[str]:
        """
        Download Piper voices concurrently with progress bar.

        :param max_workers: Maximum concurrent downloads
        :type max_workers: int
        :return: List of downloaded voice names
        :rtype: List[str]
        """
        voices_to_download: List[str] = self.list_available_voices()

        # Create semaphore to limit concurrency
        semaphore = asyncio.Semaphore(max_workers)

        async def download_with_semaphore(voice_name: str) -> str:
            async with semaphore:
                return await self.download_piper_voice(voice_name)

        tasks = [download_with_semaphore(voice) for voice in voices_to_download]

        # Progress bar with proper updates
        results = []
        with tqdm(
            total=len(tasks), desc="Downloading Piper voices", unit="voice"
        ) as pbar:
            for coro in asyncio.as_completed(tasks):
                try:
                    result = await coro
                    results.append(result)
                    pbar.update(1)
                    pbar.set_postfix({"last": result})
                except Exception as e:
                    print(f"Error downloading: {e}")
                    pbar.update(1)

        return results

    def download_piper_voices(self, max_workers: int = 3) -> List[str]:
        """
        Synchronous wrapper to download Piper voices.

        :param max_workers: Maximum concurrent downloads
        :type max_workers: int
        :return: List of downloaded voice names
        :rtype: List[str]
        """
        return asyncio.run(self._aysync_download_piper_voices(max_workers=max_workers))

    def synthesize_speech(self, text: str, language: str, output_path: Path):
        """
        Synthesize speech from text using Piper TTS.

        :param text: Text to synthesize
        :type text: str
        :param language: Language code for the voice (e.g., 'en_US')
        :type language: str
        :param output_path: Path to save the synthesized audio file
        :type output_path: Path
        """
        assert output_path.suffix in [".wav"], "Unsupported audio format"

        voice_name: str = self.get_voice(language)
        voice_path: Path = self.VOICES_DIR / f"{voice_name}.onnx"

        voice: PiperVoice = PiperVoice.load(model_path=voice_path)
        wav_file: Wave_write = Wave_write(str(output_path))
        voice.synthesize_wav(text, wav_file)


if __name__ == "__main__":
    tts_engine: TextToSpeechEngine = TextToSpeechEngine()

    voices: list[str] = tts_engine.list_available_voices()
    pprint(voices)

    languages: set[str] = tts_engine.list_available_languages()
    pprint(languages)

    voice: str = tts_engine.get_voice(language="en_GB")
    pprint(voice)

    print(tts_engine.VOICES_DIR)

    tts_engine.synthesize_speech(
        text="Hello, this is a test of the Piper text to speech synthesis.",
        language="en_GB",
        output_path=Path("test_output.wav"),
    )

    # tts_engine.download_piper_voices()
