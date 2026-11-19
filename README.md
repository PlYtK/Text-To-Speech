# Text-To-Speech

A Windows app that reads Twitch chat messages aloud using Windows SAPI5 voices.

## Features
- Reads chat messages that start with `!` (for example `!tts hello` or `!hello`)
- Voice and volume selection
- Blacklist of users that will be ignored
- Log tab with all processed messages

## Built with
- Python
- PyQt5
- TwitchIO 2.7.0
- pywin32 (Windows SAPI5)

## Requirements
- Windows
- At least one SAPI5 voice installed
- A Twitch OAuth token and your channel name

## Installation
1. Open the [Releases](../../releases) page
2. Download `TwitchTTS.exe` from the latest version
3. Run the `.exe` file

## Usage
1. Enter your Twitch token and channel name on the **Main** tab
2. Choose a voice and set the volume
3. Press **Start**
4. Send `!tts your text` in chat, or any message starting with `!`
5. To ignore a user, add their nickname on the **Blacklist** tab

## Settings
Settings are saved to `settings.json` next to the `.exe`.
If you enable "Remember token", the token is stored there as plain text,
so do not share this file.

## Run from source
```
pip install -r requirements.txt
python app.py
```
