import pyaudio
import grpc
import yandex.cloud.ai.stt.v3.stt_pb2 as stt_pb2
import yandex.cloud.ai.stt.v3.stt_service_pb2_grpc as stt_service_pb2_grpc
import pyttsx3
from es_xls import *
from win32com.client import constants as cc
import re
from words2numsrus import NumberExtractor
import pymorphy2
import gspread
import pandas as pd
import json

with open('config.json', 'r') as file:
    config = json.load(file)

lemma = pymorphy2.MorphAnalyzer()
extractor = NumberExtractor()
wb = get_excel(config.get("file_path"))
ws = wb.ActiveSheet

FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 8000
CHUNK = 4096
audio = pyaudio.PyAudio()

dawords = ['очисти', 'удали', 'прибавь', 'убавь', 'плюс', 'минус', 'заполни']


def gen():
   recognize_options = stt_pb2.StreamingOptions(
      recognition_model=stt_pb2.RecognitionModelOptions(
         audio_format=stt_pb2.AudioFormatOptions(
            raw_audio=stt_pb2.RawAudio(
               audio_encoding=stt_pb2.RawAudio.LINEAR16_PCM,
               sample_rate_hertz=8000,
               audio_channel_count=1
            )
         ),
         text_normalization=stt_pb2.TextNormalizationOptions(
            text_normalization=stt_pb2.TextNormalizationOptions.TEXT_NORMALIZATION_ENABLED,
            profanity_filter=False,
            literature_text=False
         ),
         language_restriction=stt_pb2.LanguageRestrictionOptions(
            restriction_type=stt_pb2.LanguageRestrictionOptions.WHITELIST,
            language_code=['ru-RU']
         ),
         audio_processing_type=stt_pb2.RecognitionModelOptions.REAL_TIME
      )
   )

   yield stt_pb2.StreamingRequest(session_options=recognize_options)

   stream = audio.open(format=FORMAT, channels=CHANNELS,
               rate=RATE, input=True,
               frames_per_buffer=CHUNK)
   print("recording")
   frames = []

   while True:
      data = stream.read(CHUNK)
      yield stt_pb2.StreamingRequest(chunk=stt_pb2.AudioChunk(data=data))
      frames.append(data)


class speech_bot:
    status = ""
    value = None


def is_float(string):
    try:
        float(string)
        return True
    except ValueError:
        return False


def set_tab(query):
    global ws
    sheets = {}
    for n, sh in enumerate(wb.Sheets):
        sheets[n] = sh.Name.lower()

    if 'включи лист' in query:
        s = extractor.replace_groups(query.replace('включи лист ', ''))
        if s and s.isdigit():
            wb.Sheets[int(s)].Select()
            ws = wb.ActiveSheet
            return
    if 'проект' in query:
        s = query.replace('включи ', '').replace('проект ', '')
        speech_vars = [s, s.replace(' ', ''), extractor.replace_groups(s.replace(' ', '')), extractor.replace_groups(s), extractor.replace_groups(s).replace(' ', '')]
        sheet_n = [x+1 for x in sheets if sheets[x] in speech_vars]
    if sheet_n:
        wb.Sheets[sheet_n[0]].Select()
    elif 'лист' in query or 'проект' in query:
        s = extractor.replace_groups(s.replace('лист', '').replace('включи проект', '').replace(' ', ''))
        if s and s.isdigit():
            wb.Sheets[int(s)].Select()

    ws = wb.ActiveSheet


def process_vals(vals, extractor, lemma, column_name):
    vals_str = list(vals[0])
    vals_lem = [lemma.parse(str(item))[0].normal_form if isinstance(item, str) else item for item in vals_str]
    vals_num = [extractor.replace_groups(str(item)) if not isinstance(item, str) else extractor.replace_groups(item) for item in vals_lem]
    vals_clean = [re.sub(r'[^a-zа-яA-ZА-ЯёЁ0-9]', '', str(item)) for item in vals_num]
    col_name_lem = ' '.join([lemma.parse(word)[0].normal_form for word in column_name.split()])
    return vals_clean, col_name_lem


def cell_input(query):
    # заполни строка Х столбец У ИЛИ строка Х столбец У
    pattern_1 = r'строка\s*(.*?)\s*(?=столбец|$)'
    pattern_2 = r'столбец\s*(.*?)\s*(?=строка|$)'
    row1 = re.findall(pattern_1, query)
    col1 = re.findall(pattern_2, query)

    column_name = col1[0].strip()
    ws = wb.ActiveSheet
    rn = ws.Columns(1).Find(column_name.replace(' ', ''))
    if not rn and extractor.replace_groups(column_name).isdigit():
        rn_column = extractor.replace_groups(column_name)
        rn = ws.Rows(1).Columns(int(rn_column))
        column_name = ws.Cells(1, int(rn_column)).Value
    elif not rn:
        y_last = ws.Rows(1).End(cc.xlToRight).Column
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(1, y_last)).Value
        vals_clean, col_name_lem = process_vals(vals, extractor, lemma, column_name)
        if extractor.replace_groups(col_name_lem).replace(' ', '') in vals_clean:  # убрав пробелы и заменив числа, ищу инпут в списке
            r_index = vals_clean.index(extractor.replace_groups(col_name_lem).replace(' ', ''))
            rn = ws.Columns(r_index + 1)
        else:
            speak(f"не могу найти столбец {column_name}")
            bot.status = ''
            return
    rn = rn.Column

    row_name = row1[0].strip()
    rd = ws.Columns(1).Find(row_name.replace(' ', ''))
    if not rd and extractor.replace_groups(row_name).isdigit():
        rn_row = extractor.replace_groups(row_name)
        rd = ws.Rows(int(rn_row)).Columns(1)
        row_name = ws.Cells(int(rn_row), 1).Value
    elif not rd:
        y_last = ws.Cells(ws.Rows.Count, 1).End(cc.xlUp).Row
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(y_last, 1)).Value
        vals_clean, row_name_lem = process_vals(vals, extractor, lemma, column_name)
        if extractor.replace_groups(row_name_lem).replace(' ', '') in vals_clean:
            r_index = vals_clean.index(extractor.replace_groups(row_name_lem).replace(' ', ''))
            rd = ws.Rows(r_index + 1)
        else:
            speak(f"не могу найти строку {row_name}")
            bot.status = ''
            return
    rd = rd.Row

    if rn and rd:
        speak(f"координаты: {rn} {column_name}, {rd} {row_name}")
        bot.status = "ожидание значения"
        bot.value = rn
        bot.value1 = rd
    else:
        speak(f"чёто не то. давайте ещё раз")
        bot.status = ''
        return


def value_input(query):
    ws = wb.ActiveSheet
    if "очисти" in query or "удали" in query:
        ws.Cells(bot.value1, bot.value).Value = ""
    if "прибавь" in query or "убавь" in query or "плюс" in query or "минус" in query:
        operation = '+' if 'плюс' in query or 'прибавь' in query else ''
        operation = '-' if not operation and 'минус' in query or 'убавь' in query else operation
        amount = extractor.replace_groups(query.replace('плюс', '').replace('минус', '').replace('прибавь', '').replace('убавь', '')).strip()
        if is_float(amount):
            amount = float(amount)
            cell_value = ws.Cells(bot.value1, bot.value).Value if ws.Cells(bot.value1, bot.value).Value else 0
            if operation == '+':
                ws.Cells(bot.value1, bot.value).Value = cell_value + amount
            if operation == '-':
                ws.Cells(bot.value1, bot.value).Value = cell_value - amount
        else:
            speak('Не распознал число')
    elif "заполни ячейку" in query:
        ws.Cells(bot.value1, bot.value).Value = extractor.replace_groups(query.replace('заполни ячейку ', ''))
    bot.status = ''
    bot.value = None
    bot.value1 = None


def export():
    gc = gspread.service_account()
    my_spreadsheet = gc.open("Untitled spreadsheet").sheet1     # название гугл шита
    df = pd.read_excel(config.get("file_path"))
    df.columns = ['' if isinstance(col, str) and 'Unnamed' in col else col for col in df.columns]
    df = df.apply(lambda row: row.map(lambda x: f"{x * 100}%" if isinstance(x, (int, float)) else x), axis=1)
    df = df.fillna(value='')
    my_spreadsheet.clear()
    my_spreadsheet.update([df.columns.values.tolist()] + df.values.tolist())
    speak('выгружено')


def speak(text):
    engine.say(text)
    print(text)
    engine.runAndWait()


if __name__ == '__main__':
   bot = speech_bot()
   engine = pyttsx3.init()
   voices = engine.getProperty('voices')
   engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_RU-RU_IRINA_11.0')
   cred = grpc.ssl_channel_credentials()
   channel = grpc.secure_channel('stt.api.cloud.yandex.net:443', cred)
   stub = stt_service_pb2_grpc.RecognizerStub(channel)
   it = stub.RecognizeStreaming(gen(), metadata=(('authorization', f'Api-Key {config.get("secret")}'),))

   for r in it:
      event_type, query = r.WhichOneof('Event'), None
      if event_type == 'final':
         query = [a.text for a in r.final.alternatives]
         query = query[0]
         user_words = query.split()
         matches = [word for word in user_words if word in dawords]
         if query:
            print(query)
            if "проект" in query or 'включи лист' in query:
               set_tab(query)
            if "выгружай" in query or "выгружай проект" in query:
                export()
            if "строка" in query:
                cell_input(query)
            if matches:
                value_input(query)
            if "сохрани изменения" in query:
                wb.Save()
            if "конец работы" in query:
                wb.Close()