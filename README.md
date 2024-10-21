# Presentation Voiceover
A Python tool that converts slide notes in PowerPoint presentations into voiceovers using OpenAI's text-to-speech API.

## Installation

### Prerequisites

- Python: You need to have Python installed on your machine. You can download it from the [official Python website](https://www.python.org/downloads/). After downloading, run the installer and follow the instructions. To verify the installation, open a command prompt and type `python --version`. You should see the Python version number.

- PowerPoint: Make sure you have Microsoft PowerPoint installed on your machine. If not, you can get it from the [Microsoft Office website](https://www.microsoft.com/en-us/microsoft-365/get-started-with-office-2019).

- OpenAI Account: You need an OpenAI account to use their text-to-speech API. You can create an account on the [OpenAI website](https://www.openai.com/). Once you have an account, you can find your API key in the OpenAI dashboard.

### Install Required Python Packages

Open a command prompt and type the following commands:

- Install pywin32: `pip install pywin32`
- Install openai: `pip install openai`

## Usage

1. Create a PowerPoint presentation and add your slides.
2. For each slide, add notes that you want to be transformed into audio. You can add notes in PowerPoint by clicking on the "Notes" section at the bottom of the slide and typing your text.
3. Save your presentation.
4. Clone the GitHub repository to your local machine: `git clone https://github.com/maksym-odanets/presentation-voiceover`
5. Navigate to the directory: `cd presentation-voiceover/src`
6. Run the script with the necessary arguments. For example: `python main.py --api-key YOUR_OPENAI_API_KEY --pptx-file PATH_TO_YOUR_PPTX_FILE`

Replace `YOUR_OPENAI_API_KEY` with your actual OpenAI API key, `PATH_TO_YOUR_PPTX_FILE` with the path to the PowerPoint file you want to add voiceovers to.

The script will generate voiceovers for the slide notes in your PowerPoint presentation and save a new presentation with the voiceovers embedded. The audio files will be stored in the specified audio directory, or in a new directory in the system's temp directory if no audio directory is specified or if the specified directory does not exist.

### Audio File Caching

For cost-effectiveness, all audio files (voiceovers) are cached. This means that if you run the script multiple times with the same arguments, it will not regenerate the voiceovers unless you use the `--no-cache` or `--slides` arguments to clear the cache.

### Regenerating All Slides

If you want to regenerate the voiceovers for all slides, you can use the `--no-cache` argument:

`python main.py --api-key YOUR_OPENAI_API_KEY --pptx-file PATH_TO_YOUR_PPTX_FILE --no-cache`

This will clear the audio cache and regenerate the voiceovers for all slides.

### Regenerating Specific Slides

If you want to regenerate the voiceovers for specific slides, you can use the `--slides` argument:

`python main.py --api-key YOUR_OPENAI_API_KEY --pptx-file PATH_TO_YOUR_PPTX_FILE --slides "1,3,5"`

This will regenerate the voiceovers for slides 1, 3, and 5.

### Arguments

- `--api-key`: Required. Your OpenAI API key.
- `--pptx-file`: Required. The path to the PowerPoint presentation you want to add voiceovers to.
- `--model`: Optional. The model to use for text-to-speech generation. Default is "tts-1".
- `--voice`: Optional. The voice to use for text-to-speech generation. Default is "alloy". You can find other possible voices on the [OpenAI Text-to-Speech guide](https://platform.openai.com/docs/guides/text-to-speech).
- `--audio-dir`: Optional. The directory where you want to store the audio files. If not provided or if the directory does not exist, a new directory will be created in the system's temp directory with the name of the presentation.
- `--no-cache`: Optional. If set, the audio directory will be emptied before processing.
- `--slides`: Optional. A comma-separated list of slide numbers to regenerate. If provided, only the audio files for these slides will be regenerated.

## Known Issues

- After adding a sound to a slide, the existing Animation might be reset to be a Default "Appear". This is a known issue and we are looking into ways to preserve the original animation settings when adding a voiceover. In the meantime, you may need to manually reset the animation settings after running the script. We apologize for any inconvenience this may cause.

## Buy me a coffee
If you find my content helpful and would like to show your appreciation, please consider buying me a coffee. Your support is greatly appreciated! 😊

[![Buy Me a Coffee](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/donate/?hosted_button_id=3SAZN958APPKW)