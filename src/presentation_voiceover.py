from win32com.client import Dispatch

from logger import LoggerInterface
from speech_generator import SpeechGeneratorInterface

class PresentationVoiceover:
    def __init__(self, logger: LoggerInterface, tts: SpeechGeneratorInterface):
        self.logger = logger
        self.tts = tts

    def handle(self, prs: Dispatch, audio_dir: str):
        self.audio_dir = audio_dir

        slide_width = prs.PageSetup.SlideWidth
        left_position = slide_width + 10

        for slide_number, slide in enumerate(prs.Slides, start=1):
            self.process_slide(slide, slide_number, left_position)

    def process_slide(self, slide, slide_number: int, left_position: int):
        notes_text = self.extract_notes(slide, slide_number)
        if not notes_text:
            return

        speech_file_path = self.audio_dir / f"slide_{slide_number}_speech.mp3"
        if not speech_file_path.exists():
            if not self.tts.generate_speech(notes_text, speech_file_path):
                self.logger.error(f"Failed to generate speech for slide {slide_number}")
                return

        self.embed_audio(slide, speech_file_path, left_position)

    def extract_notes(self, slide, slide_number: int):
        notes_page = slide.NotesPage
        if notes_page.Shapes.Count >= 2:
            notes_shape = notes_page.Shapes.Item(2)
            return notes_shape.TextFrame.TextRange.Text
        else:
            self.logger.info(f"Slide {slide_number} does not have a second shape for notes.")
            return ""

    def embed_audio(self, slide, audio_path: str, left_position: int):
        try:
            audio_path_str = str(audio_path)
            audio_shape = slide.Shapes.AddMediaObject2(FileName=audio_path_str, LinkToFile=False, SaveWithDocument=True, Left=left_position, Top=0)
            self.configure_audio_settings(audio_shape)
            self.logger.info(f"Embedded audio into slide: {audio_path_str}")
        except Exception as e:
            self.logger.error(f"Failed to embed audio: {e}")

    def configure_audio_settings(self, audio_shape):
        anim_settings = audio_shape.AnimationSettings
        play_settings = anim_settings.PlaySettings
        play_settings.PlayOnEntry = True
        play_settings.HideWhileNotPlaying = False
        play_settings.LoopUntilStopped = False
        play_settings.StopAfterSlides = 1