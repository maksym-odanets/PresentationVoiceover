import argparse
import win32com.client
import tempfile
import os

from pathlib import Path
from logger import ConsoleLogger
from speech_generator import OpenAISpeechGenerator
from presentation_voiceover import PresentationVoiceover

def parse_arguments():
    parser = argparse.ArgumentParser(description="PowerPoint Voiceover.")
    parser.add_argument("--api-key", required=True, help="API key to access the OpenAI")
    parser.add_argument("--pptx-file", required=True, help="The path to the PowerPoint presentation")
    parser.add_argument("--audio-dir", help="The directory to store audio files")
    parser.add_argument("--model", default="tts-1", help="Model to use for text-to-speech generation")
    parser.add_argument("--voice", default="alloy", help="Voice to use for text-to-speech generation")
    return parser.parse_args()

def main():
    args = parse_arguments()

    logger = ConsoleLogger()
    tts = OpenAISpeechGenerator(args.api_key, args.model, args.voice)

    app = win32com.client.Dispatch("PowerPoint.Application")
    prs = app.Presentations.Open(args.pptx_file)

    audio_dir = args.audio_dir
    if not audio_dir or not os.path.exists(audio_dir):
        presentation_name = Path(args.pptx_file).stem
        audio_dir = Path(tempfile.gettempdir()) / presentation_name
    os.makedirs(audio_dir, exist_ok=True)

    pptx_notes_to_voiceover = PresentationVoiceover(logger, tts)
    pptx_notes_to_voiceover.handle(prs, audio_dir)

    modified_pptx_file = args.pptx_file.replace('.pptx', '_modified.pptx')
    prs.SaveAs(modified_pptx_file)
    prs.Close()
    app.Quit()

    logger.info(f"Saved modified presentation to {modified_pptx_file}")

if __name__ == "__main__":
    main()