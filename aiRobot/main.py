import os
import time
import random
import winsound
import win32com.client
from aip import AipSpeech
import speech_recognition as sr


BAIDU_APP_ID = 'https://openapi.baidu.com/oauth/2.0/token?'
BAIDU_API_KEY = 'vyYLov63W6x33nIPvwVdLLsX'
BAIDU_SECRET_KEY = 'EzgQoFF9xp62SeGCXCBaD8FjWNxvl9kZ'
API_Input_Speech = AipSpeech(BAIDU_APP_ID, BAIDU_API_KEY, BAIDU_SECRET_KEY)

API_Output_Speech = win32com.client.Dispatch("SAPI.SPVOICE")


# itemBank Local
fileLoc_CN = "itemBank/computer_Network/items.txt"
fileLoc_DS = "itemBank/data_Structure/items.txt"
fileLoc_DB = "itemBank/dataBase/items.txt"
fileLoc_LI = "itemBank/language_Infrastructure/c++.txt"
fileLoc_OS = "itemBank/operating_System/items.txt"


def textToVoice(text_data):
    API_Output_Speech.Speak(text_data)
    winsound.PlaySound(text_data, winsound.SND_ASYNC)


def voiceToText(audio_data):
    result = API_Input_Speech.asr(audio_data, 'wav', 16000, {'dev_pid': 1536})
    try:
        text = result['result'][0]
    except Exception as e:
        print(e)
        text = ""
    return text


def get_Message():
    # Initializing Recognizer
    r = sr.Recognizer()

    # Initializing the Microphone
    mic = sr.Microphone(sample_rate=16000)

    # Recording
    print("开始")
    with mic as source:
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    # Analysis
    audio_data = audio.get_wav_data()

    return voiceToText(audio_data)


def language_Infrastructure():
    with open(fileLoc_LI, mode='r', encoding='utf-8') as f:
        itemLines = f.read().splitlines()

    text = ""
    textToVoice("开始")
    while True:
        message = get_Message()
        if message == "返回主菜单":
            break
        elif message == "下一题":
            text = random.choice(itemLines)
            textToVoice(text)
        elif message == "再说一遍题目":
            textToVoice(text)
        else:
            textToVoice("听不清，请再说一遍吧")


def data_Structure():
    with open(fileLoc_DS, mode='r', encoding='utf-8') as f:
        itemLines = f.read().splitlines()

    text = ""
    textToVoice("开始")
    while True:
        message = get_Message()
        if message == "返回主菜单":
            break
        elif message == "下一题":
            text = random.choice(itemLines)
            textToVoice(text)
        elif message == "再说一遍题目":
            textToVoice(text)
        else:
            textToVoice("听不清，请再说一遍吧")


def operating_System():
    with open(fileLoc_OS, mode='r', encoding='utf-8') as f:
        itemLines = f.read().splitlines()

    text = ""
    textToVoice("开始")
    while True:
        message = get_Message()
        if message == "返回主菜单":
            break
        elif message == "下一题":
            text = random.choice(itemLines)
            textToVoice(text)
        elif message == "再说一遍题目":
            textToVoice(text)
        else:
            textToVoice("听不清，请再说一遍吧")


def computer_Network():
    with open(fileLoc_CN, mode='r', encoding='utf-8') as f:
        itemLines = f.read().splitlines()

    text = ""
    textToVoice("开始")
    while True:
        message = get_Message()
        if message == "返回主菜单":
            break
        elif message == "下一题":
            text = random.choice(itemLines)
            textToVoice(text)
        elif message == "再说一遍题目":
            textToVoice(text)
        else:
            textToVoice("听不清，请再说一遍吧")


def dataBase():
    with open(fileLoc_DB, mode='r', encoding='utf-8') as f:
        itemLines = f.read().splitlines()

    text = ""
    textToVoice("开始")
    while True:
        message = get_Message()
        if message == "返回主菜单":
            break
        elif message == "下一题":
            text = random.choice(itemLines)
            textToVoice(text)
        elif message == "再说一遍题目":
            textToVoice(text)
        else:
            textToVoice("听不清，请再说一遍吧")


def main():
    while True:
        textToVoice("请选择你要学习的内容")
        print("\n请选择你要学习的内容")
        select_itemBank = get_Message()
        os.system("cls")
        if select_itemBank == "关机":
            break

        elif select_itemBank == "语法基础":
            textToVoice("已成功进入语法基础的学习")
            language_Infrastructure()

        elif select_itemBank == "数据结构":
            textToVoice("已成功进入数据结构的学习")
            data_Structure()

        elif select_itemBank == "操作系统":
            textToVoice("已成功进入操作系统的学习")
            operating_System()

        elif select_itemBank == "计算机网络":
            textToVoice("已成功进入计算机网络的学习")
            computer_Network()

        elif select_itemBank == "数据库":
            textToVoice("已成功进入数据库的学习")
            dataBase()

        else:
            textToVoice("听不清，请再说一遍吧")


if __name__ == '__main__':
    main()
    textToVoice("已成功关机")
