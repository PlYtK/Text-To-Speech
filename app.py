import asyncio
import json
import os
import queue
import sys
import threading
import time
from datetime import datetime
from xml.sax.saxutils import escape

import pythoncom
import win32com.client
from PyQt5 import uic
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QMainWindow
from twitchio.ext import commands

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UI_PATH = os.path.join(BASE_DIR, 'MainWindow.ui')
SETTINGS_PATH = os.path.join(BASE_DIR, 'settings.json')

BLACKLIST = {'idktion'}   # ники в нижнем регистре
RATE = 0                  # от -10 до 10
PAUSE_MS = 300            # тишина перед фразой
TAIL_SEC = 0.3            # ожидание после фразы


def list_voices():
    pythoncom.CoInitialize()
    sapi = win32com.client.Dispatch('SAPI.SpVoice')
    return [v.GetDescription() for v in sapi.GetVoices()]


class TTS:
    def __init__(self):
        self.q = queue.Queue()
        self.voice_name = ''
        self.volume = 25
        threading.Thread(target=self._worker, daemon=True).start()

    def speak(self, text):
        text = text.replace('<', ' меньше чем ').strip()
        if not text:
            return
        if text[-1] not in '.!?…':
            text += '.'
        self.q.put(text)

    def _worker(self):
        pythoncom.CoInitialize()
        sapi = win32com.client.Dispatch('SAPI.SpVoice')
        current = None
        while True:
            text = self.q.get()
            try:
                if self.voice_name and self.voice_name != current:
                    for v in sapi.GetVoices():
                        if v.GetDescription() == self.voice_name:
                            sapi.Voice = v
                            current = self.voice_name
                            break
                sapi.Volume = self.volume
                sapi.Rate = RATE
                sapi.Speak(f'<silence msec="{PAUSE_MS}"/>{escape(text)}', 8)
                time.sleep(TAIL_SEC)
            except Exception as e:
                print(f'TTS error: {e}')


class TwitchBot(commands.Bot):
    def __init__(self, token, channel, loop, on_log, on_text):
        super().__init__(token=token, prefix='!!!',
                         initial_channels=[channel], loop=loop)
        self.on_log = on_log
        self.on_text = on_text

    async def event_ready(self):
        self.on_log(f'Подключено как {self.nick}')

    async def event_message(self, message):
        if message.echo or not message.content:
            return
        if message.author.name.lower() in BLACKLIST:
            return

        content = message.content
        text = None
        if content.startswith(('!tts ', '!ттс ')):
            text = content[5:]
        elif content.startswith('!'):
            text = content[1:]

        if text and text.strip():
            self.on_log(f'{message.author.name}: {text}')
            self.on_text(text)


class BotThread(QThread):
    log = pyqtSignal(str)
    stopped = pyqtSignal()

    def __init__(self, token, channel, tts):
        super().__init__()
        self.token = token.replace('oauth:', '').strip()
        self.channel = channel.strip().lstrip('#')
        self.tts = tts
        self.loop = None
        self.bot = None

    def run(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        try:
            self.bot = TwitchBot(self.token, self.channel, self.loop,
                                 self.log.emit, self.tts.speak)
            self.loop.run_until_complete(self.bot.start())
        except Exception as e:
            self.log.emit(f'Ошибка: {e}')
        finally:
            self.loop.close()
            self.stopped.emit()

    def stop(self):
        if self.bot and self.loop and self.loop.is_running():
            asyncio.run_coroutine_threadsafe(self.bot.close(), self.loop)


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(UI_PATH, self)

        self.tts = TTS()
        self.thread = None

        self.voiceBox.addItems(list_voices())
        self.load_settings()
        self.apply_audio()

        self.voiceBox.currentTextChanged.connect(self.apply_audio)
        self.volumeSlider.valueChanged.connect(self.apply_audio)
        self.startButton.clicked.connect(self.toggle)
        self.clearLogButton.clicked.connect(self.logView.clear)

    def apply_audio(self, *_):
        self.tts.voice_name = self.voiceBox.currentText()
        self.tts.volume = self.volumeSlider.value()

    def append_log(self, text):
        ts = datetime.now().strftime('%H:%M:%S')
        self.logView.appendPlainText(f'[{ts}] {text}')
        if self.autoScrollCheck.isChecked():
            sb = self.logView.verticalScrollBar()
            sb.setValue(sb.maximum())

    def toggle(self):
        if self.thread and self.thread.isRunning():
            self.startButton.setEnabled(False)
            self.thread.stop()
            return

        token = self.tokenEdit.text().strip()
        channel = self.channelEdit.text().strip()
        if not token or not channel:
            self.append_log('Введите токен и имя канала')
            self.tabWidget.setCurrentWidget(self.tabLog)
            return

        self.thread = BotThread(token, channel, self.tts)
        self.thread.log.connect(self.append_log)
        self.thread.stopped.connect(self.on_stopped)
        self.thread.start()
        self.startButton.setText('Остановить')
        self.append_log('Запуск...')

    def on_stopped(self):
        self.startButton.setText('Запустить')
        self.startButton.setEnabled(True)
        self.append_log('Остановлено')

    def load_settings(self):
        if not os.path.exists(SETTINGS_PATH):
            return
        try:
            with open(SETTINGS_PATH, encoding='utf-8') as f:
                s = json.load(f)
        except Exception:
            return
        self.channelEdit.setText(s.get('channel', ''))
        token = s.get('token', '')
        self.tokenEdit.setText(token)
        self.rememberCheck.setChecked(bool(token))
        self.volumeSlider.setValue(s.get('volume', 25))
        i = self.voiceBox.findText(s.get('voice', ''))
        if i >= 0:
            self.voiceBox.setCurrentIndex(i)

    def closeEvent(self, event):
        s = {
            'channel': self.channelEdit.text(),
            'voice': self.voiceBox.currentText(),
            'volume': self.volumeSlider.value(),
            'token': self.tokenEdit.text() if self.rememberCheck.isChecked() else '',
        }
        with open(SETTINGS_PATH, 'w', encoding='utf-8') as f:
            json.dump(s, f, ensure_ascii=False, indent=2)

        if self.thread and self.thread.isRunning():
            self.thread.stop()
            self.thread.wait(3000)
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Window()
    w.show()
    sys.exit(app.exec_())