#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@File:        utils.py
@Author:      lgx
@Contact:     1101588023@qq.com
@Time:        2024/01/05 10:48
@Description:
"""

import json
import os
from pprint import pprint
from uuid import uuid4
import asyncio
import winsound
from pathlib import Path
import edge_tts
import ffmpeg
from waapi import WaapiClient
from openpyxl import load_workbook

VO_LIST_FILE = Path('vo_list.json')


# TEMPL_PATH = Path('templ')
# if not TEMPL_PATH.is_dir():
#     TEMPL_PATH.mkdir()


async def get_vo_list():
    vo_list = await edge_tts.list_voices()
    json_str = json.dumps(vo_list, indent=2)
    VO_LIST_FILE.write_text(json_str, encoding='utf-8')
    return vo_list


async def load_vo_list():
    if not VO_LIST_FILE.is_file():
        return await get_vo_list()
    else:
        json_str = VO_LIST_FILE.read_text(encoding='utf-8')
        return json.loads(json_str)


async def play_back(speaker: str, text: str):
    if not speaker or not text: return
    text = text.splitlines()[0].strip()
    filepath = 'tts_templ.mp3'
    communicate = edge_tts.Communicate(text, speaker)
    await communicate.save(str(filepath))
    os.startfile(filepath)


async def save_audio(speaker: str, text: str, filepath: str):
    await synthesis(speaker, text, filepath)
    os.startfile(Path(filepath).parent)


async def synthesis(speaker: str, text: str, filepath: str):
    if not speaker or not text: return
    communicate = edge_tts.Communicate(text, speaker)

    if filepath.endswith('.wav'):
        temp_path = filepath + '.mp3'
        await communicate.save(temp_path)
        ffmpeg_convert(temp_path, filepath)
        return

    if not filepath.endswith('.mp3'):
        filepath += '.mp3'
    await communicate.save(filepath)


def ffmpeg_convert(orig_file: str | Path, target_file: str | Path):
    if isinstance(orig_file, str):
        orig_file = Path(orig_file)
    if isinstance(target_file, str):
        target_file = Path(target_file)

    if not orig_file.is_file():
        raise FileNotFoundError(f'{orig_file}')

    if not target_file.is_file():
        ffmpeg.input(str(orig_file)).output(str(target_file)).run()
        assert target_file.is_file()
        orig_file.unlink()
    else:
        templ_file = target_file.with_stem(str(uuid4()))
        ffmpeg.input(str(orig_file)).output(str(templ_file)).run()
        assert templ_file.is_file()
        target_file.unlink()
        templ_file.rename(target_file.name)
        orig_file.unlink()


def play_sound(path: str):
    winsound.PlaySound(path, winsound.SND_FILENAME)


def waapi_get_lang_list():
    args = {
        'from': {'ofType': ['Language']}
    }
    pass_lang = ['Mixed', 'SFX', 'External']
    with WaapiClient(allow_exception=True) as client:
        result = client.call('ak.wwise.core.object.get', args)
        if result and result['return']:
            return [i['name'] for i in result['return'] if i['name'] not in pass_lang]


def waapi_get_selected():
    with WaapiClient(allow_exception=True) as client:
        result = client.call('ak.wwise.ui.getSelectedObjects', options={'return': ['path']})
        if result and result['objects']:
            return result['objects'][0]['path']


def waapi_import_vo(data: list):
    imports = []
    sel_path: str = waapi_get_selected()
    if not sel_path or not sel_path.startswith('\\Actor-Mixer Hierarchy'):
        raise Exception(f'Error parent path: {sel_path}')

    for (name, lang, speaker, text, filepath) in data:
        imports.append({
            'audioFile': str(filepath),
            'objectPath': sel_path + f'\\<Sound Voice>{name}' + f'\\<AudioFileSource>{filepath.stem}',
            'importLanguage': lang
        })

    # Import the generated wav files to Wwise
    import_args = {
        'importOperation': 'useExisting',
        'imports': imports,
        'autoAddToSourceControl': True
    }
    with WaapiClient(allow_exception=False) as client:
        client.call('ak.wwise.core.audio.import', import_args)


def load_xl(file: str):
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    data = []
    headers = {'VoiceName': -1, 'Language': -1, 'Speaker': -1, 'Text': -1}
    # get header index
    for i in ws.values:
        for col, j in enumerate(i):
            if j in headers:
                headers[j] = col
        break

    is_first = True
    for i in ws.values:
        if is_first:
            is_first = False
            continue
        row_data = {}
        for k, v in headers.items():
            if v != -1:
                row_data[k] = i[v]
        data.append(row_data)

    return data


if __name__ == '__main__':
    # asyncio.run(get_vo_list())
    # print(asyncio.run(load_vo_list()))
    # play_sound(r'C:\Users\Administrator\Music\MONACA - 遺サレタ場所_斜光.wav')
    # ffmpeg_convert('1.mp3', '2.wav')
    # waapi_get_lang_list()
    # print(waapi_get_selected())
    load_xl(r'C:\Users\Administrator\Downloads\新建 XLSX 工作表.xlsx')
