from openai import OpenAI

class SpeechGeneratorInterface:
    def generate_speech(self, text: str, speech_file_path: str):
        pass

class OpenAISpeechGenerator(SpeechGeneratorInterface):
    def __init__(self, api_key: str, model: str, voice: str):
        self.client = OpenAI(api_key=api_key)
        self.model = model
        self.voice = voice

    def generate_speech(self, text: str, speech_file_path: str):
        try:
            response = self.client.audio.speech.create(model=self.model, voice=self.voice, input=text)
            with open(speech_file_path, 'wb') as f:
                f.write(response.content)
            return True
        except Exception as e:
            return False