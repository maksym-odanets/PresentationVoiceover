import hashlib
from pathlib import Path
from abc import ABC, abstractmethod
from typing import Optional

from openai import OpenAI


class AbstractSpeechGenerator(ABC):
    def generate_speech(self, text: str, speech_file_path: Path) -> bool:
        if speech_file_path.exists():
            return True

        return self._generate_speech_from_api(text, speech_file_path)

    @abstractmethod
    def _generate_speech_from_api(self, text: str, speech_file_path: Path) -> bool:
        pass

    def clear_cache(self, audio_dir: Path, text: Optional[str] = None) -> None:
        if text:
            audio_file = audio_dir / self.get_audio_filename(text)
            if audio_file.exists():
                audio_file.unlink()
        else:
            for audio_file in audio_dir.glob("*.mp3"):
                audio_file.unlink()

    def get_audio_filename(self, text: str) -> str:
        notes_hash = hashlib.md5(text.encode()).hexdigest()
        return f"{notes_hash}.mp3"


class OpenAISpeechGenerator(AbstractSpeechGenerator):
    def __init__(self, api_key: str, model: str, voice: str) -> None:
        self.client = OpenAI(api_key=api_key)
        self.model = model
        self.voice = voice

    def _generate_speech_from_api(self, text: str, speech_file_path: Path) -> bool:
        try:
            response = self.client.audio.speech.create(model=self.model, voice=self.voice, input=text)
            with open(speech_file_path, 'wb') as f:
                f.write(response.content)
            return True
        except Exception as e:
            return False