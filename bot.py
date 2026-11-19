import os
import queue
import threading
import time
from xml.sax.saxutils import escape

import pythoncom
import win32com.client
from twitchio.ext import commands

BLACKLIST = {'idktion'}          # ники в нижнем регистре
VOICE_NAME = 'Tatyana'           # часть названия голоса
VOLUME = 25                      # 0-100
RATE = 0                         # от -10 до 10, 0 = стандартная скорость
PAUSE_MS = 300                   # тишина перед фразой
TAIL_SEC = 0.3                   # ожидание после фразы

tts_queue = queue.Queue()


def tts_worker():
    pythoncom.CoInitialize()
    voice = win32com.client.Dispatch('SAPI.SpVoice')

    for v in voice.GetVoices():
        if VOICE_NAME.lower() in v.GetDescription().lower():
            voice.Voice = v
            break
    else:
        print('Голос не найден, используется голос по умолчанию')

    voice.Volume = VOLUME
    voice.Rate = RATE

    while True:
        text = tts_queue.get()
        try:
            xml = f'<silence msec="{PAUSE_MS}"/>{escape(text)}'
            voice.Speak(xml, 8)   # 8 = SVSFIsXML, вызов синхронный
            time.sleep(TAIL_SEC)
        except Exception as e:
            print(f'TTS error: {e}')
        finally:
            tts_queue.task_done()


threading.Thread(target=tts_worker, daemon=True).start()


def speak(text: str):
    text = text.replace('<', ' меньше чем ').strip()
    if not text:
        return
    if text[-1] not in '.!?…':
        text += '.'
    tts_queue.put(text)


class Bot(commands.Bot):

    def __init__(self):
        super().__init__(
            token='enter your token',
            prefix='!!!',
            initial_channels=['aklauma'],
        )

    async def event_ready(self):
        print(f'Logged in as | {self.nick}')
        print('UwU')

    async def event_message(self, message):
        if message.echo or not message.content:
            return

        if message.author.name.lower() not in BLACKLIST:
            content = message.content
            text = None

            if content.startswith(('!tts ', '!ттс ')):
                text = content[5:]
            elif content.startswith('!'):
                text = content[1:]

            if text and text.strip():
                print(f'{message.author.name}: {text}')
                speak(text)

        await self.handle_commands(message)

    @commands.command(name='qq')
    async def sr(self, ctx: commands.Context):
        pass


bot = Bot()
bot.run()