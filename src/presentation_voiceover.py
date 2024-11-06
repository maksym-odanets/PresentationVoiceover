from pathlib import Path
from typing import List, Optional

from win32com.client import Dispatch

from logger import AbstractLogger
from speech_generator import AbstractSpeechGenerator


class PresentationVoiceover:
    def __init__(self, logger: AbstractLogger, tts: AbstractSpeechGenerator) -> None:
        self.logger = logger
        self.tts = tts

    def handle(self, prs: Dispatch, audio_dir: Path, slides_to_process: Optional[List[int]] = None) -> None:
        slide_width = prs.PageSetup.SlideWidth
        left_position = slide_width + 10

        if slides_to_process:
            self._process_selected_slides(prs, audio_dir, slides_to_process, left_position)
        else:
            self._process_all_slides(prs, audio_dir, left_position)

    def _process_selected_slides(self, prs: Dispatch, audio_dir: Path, slides_to_process: List[int], left_position: int) -> None:
        for slide_number in slides_to_process:
            slide = prs.Slides(slide_number)
            notes_text = self.extract_notes(slide, slide_number)
            if notes_text:
                self.tts.clear_cache(audio_dir, notes_text)
                self.process_slide(slide, slide_number, left_position, audio_dir)

    def _process_all_slides(self, prs: Dispatch, audio_dir: Path, left_position: int) -> None:
        for slide_number, slide in enumerate(prs.Slides, start=1):
            notes_text = self.extract_notes(slide, slide_number)
            if notes_text:
                speech_file_path = audio_dir / self.tts.get_audio_filename(notes_text)
                if not speech_file_path.exists():
                    self.process_slide(slide, slide_number, left_position, audio_dir)

    def process_slide(self, slide: Dispatch, slide_number: int, left_position: int, audio_dir: Path) -> None:
        notes_text = self.extract_notes(slide, slide_number)
        if not notes_text:
            return

        speech_file_path = audio_dir / self.tts.get_audio_filename(notes_text)
        if not self.tts.generate_speech(notes_text, speech_file_path):
            self.logger.error(f"Failed to generate speech for slide {slide_number}")
            return

        self.embed_audio(slide, speech_file_path, left_position)

    def extract_notes(self, slide: Dispatch, slide_number: int) -> str:
        notes_page = slide.NotesPage
        if notes_page.Shapes.Count >= 2:
            notes_shape = notes_page.Shapes.Item(2)
            return notes_shape.TextFrame.TextRange.Text
        else:
            self.logger.info(f"Slide {slide_number} does not have a second shape for notes.")
            return ""

    def embed_audio(self, slide: Dispatch, audio_path: Path, left_position: int) -> None:
        try:
            audio_shape = slide.Shapes.AddMediaObject2(
                FileName=str(audio_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left_position,
                Top=0
            )
            self.configure_audio_settings(audio_shape)
            self.logger.info(f"Embedded audio into slide: {audio_path}")
        except Exception as e:
            self.logger.error(f"Failed to embed audio: {e}")

    def configure_audio_settings(self, audio_shape: Dispatch) -> None:
        anim_settings = audio_shape.AnimationSettings
        play_settings = anim_settings.PlaySettings
        play_settings.PlayOnEntry = True
        play_settings.HideWhileNotPlaying = False
        play_settings.LoopUntilStopped = False
        play_settings.StopAfterSlides = 1