import pyttsx3
import speech_recognition as sr
from datetime import date
import time
import webbrowser
import datetime
from pynput.keyboard import Key, Controller
import sys
import os
from os import listdir
from os.path import isfile, join
from threading import Thread
import app
from network_tracker import run_network_usage_tracker

# New Imports
import subprocess
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER

# -------------Object Initialization---------------
today = date.today()
r = sr.Recognizer()
keyboard = Controller()
engine = pyttsx3.init('sapi5')
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Audio Volume Control
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume_control = cast(interface, POINTER(IAudioEndpointVolume))

# ----------------Variables------------------------
file_exp_status = False
files = []
path = ''
is_awake = True

# ------------------Functions----------------------
def reply(audio):
    app.ChatBot.addAppMsg(audio)
    print(audio)
    engine.say(audio)
    engine.runAndWait()

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        reply("Good Morning!")
    elif hour >= 12 and hour < 18:
        reply("Good Afternoon!")
    else:
        reply("Good Evening!")
    reply("I am Nova, how may I help you?")

def run_translator_gui():
    try:
        print("Launching Voice Translator GUI...")
        import main
        main.run_translator_gui()
    except Exception as e:
        print(f"Error launching Translator: {str(e)}")
        reply("Sorry, there was an error while launching the Voice Translator.")

with sr.Microphone() as source:
    r.energy_threshold = 500
    r.dynamic_energy_threshold = False

def record_audio():
    with sr.Microphone() as source:
        r.pause_threshold = 0.8
        voice_data = ''
        audio = r.listen(source, phrase_time_limit=5)
        try:
            voice_data = r.recognize_google(audio)
        except sr.RequestError:
            reply('Sorry my service is down. Please check your internet connection.')
        except sr.UnknownValueError:
            print('Cannot recognize')
        return voice_data.lower()

# New: Volume Control
def change_volume(up=True):
    current_volume = volume_control.GetMasterVolumeLevelScalar()
    step = 0.1
    new_volume = current_volume + step if up else current_volume - step
    new_volume = max(0.0, min(1.0, new_volume))
    volume_control.SetMasterVolumeLevelScalar(new_volume, None)
    reply(f"Volume {'increased' if up else 'decreased'}")


# New: Open Windows Settings
def open_settings_page(setting):
    try:
        subprocess.run(["start", f"ms-settings:{setting}"], shell=True)
        reply(f"Opening {setting.replace('-', ' ')} settings.")
    except Exception as e:
        reply(f"Failed to open {setting} settings.")

import subprocess

# Function to open applications based on the user's command
def open_application(app_name):
    try:
        if app_name == 'notepad':
            subprocess.Popen('notepad.exe')
            reply("Opening Notepad...")
        elif app_name == 'calculator':
            subprocess.Popen('calc.exe')
            reply("Opening Calculator...")
        elif app_name == 'word':
            subprocess.Popen('winword.exe')
            reply("Opening Microsoft Word...")
        elif app_name == 'excel':
            subprocess.Popen('excel.exe')
            reply("Opening Microsoft Excel...")
        elif app_name == 'powerpoint':
            subprocess.Popen('powerpnt.exe')
            reply("Opening Microsoft PowerPoint...")
        elif app_name == 'chrome':
            subprocess.Popen('chrome.exe')
            reply("Opening Google Chrome...")
        elif app_name == 'firefox':
            subprocess.Popen('firefox.exe')
            reply("Opening Mozilla Firefox...")
        elif app_name == 'edge':
            subprocess.Popen('msedge.exe')
            reply("Opening Microsoft Edge...")
        elif app_name == 'explorer':
            subprocess.Popen('explorer.exe')
            reply("Opening File Explorer...")
        else:
            reply(f"Sorry, I can't open {app_name} yet.")
    except Exception as e:
        reply(f"Error opening {app_name}: {str(e)}")


def respond(voice_data):
    global file_exp_status, files, is_awake, path
    print(voice_data)
    voice_data = voice_data.lower().replace('nova', '')  # Ensure voice_data is properly sanitized
    app.eel.addUserMsg(voice_data)

    if not is_awake:
        if 'wake up' in voice_data:
            is_awake = True
            wish()
        return

    # STATIC CONTROLS
    if 'hello' in voice_data:
        wish()

    elif 'what is your name' in voice_data:
        reply('My name is nova!')

    elif 'date' in voice_data:
        reply(today.strftime("%B %d, %Y"))

    elif 'time' in voice_data:
        reply(str(datetime.datetime.now()).split(" ")[1].split('.')[0])

    elif 'search' in voice_data:
        reply('Searching for ' + voice_data.split('search')[1])
        url = 'https://google.com/search?q=' + voice_data.split('search')[1]
        try:
            webbrowser.get().open(url)
            reply('This is what I found')
        except:
            reply('Please check your Internet Connection')

    elif 'location' in voice_data:
        reply('Which place are you looking for?')
        temp_audio = record_audio()
        app.eel.addUserMsg(temp_audio)
        reply('Locating...')
        url = 'https://google.nl/maps/place/' + temp_audio + '/&amp;'
        try:
            webbrowser.get().open(url)
            reply('This is what I found')
        except:
            reply('Please check your Internet Connection')

    elif 'bye' in voice_data or 'by' in voice_data:
        reply("Goodbye! Have a nice day.")
        is_awake = False

    # SYSTEM SETTINGS CONTROLS
    elif 'increase volume' in voice_data:
        change_volume(up=True)

    elif 'decrease volume' in voice_data:
        change_volume(up=False)

    elif 'open hotspot' in voice_data:
        open_settings_page("network-mobilehotspot")

    elif 'open airplane mode' in voice_data or 'airplane mode' in voice_data:
        open_settings_page("network-airplanemode")

    # DYNAMIC CONTROLS
    elif 'launch translator' in voice_data or 'open translator' in voice_data:
        reply('Launching Voice Translator...')
        t2 = Thread(target=run_translator_gui)
        t2.start()

    elif 'open network usage' in voice_data or 'network usage' in voice_data:
        reply('Opening Network Usage Tracker...')
        t3 = Thread(target=run_network_usage_tracker)
        t3.start()

    elif 'copy' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('c')
            keyboard.release('c')
        reply('Copied')

    elif 'paste' in voice_data or 'pest' in voice_data or 'page' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('v')
            keyboard.release('v')
        reply('Pasted')

    elif 'list' in voice_data:
        counter = 0
        path = 'C://'
        files = listdir(path)
        filestr = ""
        for f in files:
            counter += 1
            print(str(counter) + ':  ' + f)
            filestr += str(counter) + ':  ' + f + '<br>'
        file_exp_status = True
        reply('These are the files in your root directory')
        app.ChatBot.addAppMsg(filestr)

    elif file_exp_status:
        counter = 0
        if 'open' in voice_data:
            index = int(voice_data.split(' ')[-1]) - 1
            if isfile(join(path, files[index])):
                os.startfile(path + files[index])
                file_exp_status = False
            else:
                try:
                    path += files[index] + '//'
                    files = listdir(path)
                    filestr = ""
                    for f in files:
                        counter += 1
                        filestr += str(counter) + ':  ' + f + '<br>'
                        print(str(counter) + ':  ' + f)
                    reply('Opened Successfully')
                    app.ChatBot.addAppMsg(filestr)
                except:
                    reply('You do not have permission to access this folder')

        elif 'back' in voice_data:
            filestr = ""
            if path == 'C://':
                reply('Sorry, this is the root directory')
            else:
                path = '//'.join(path.split('//')[:-2]) + '//'
                files = listdir(path)
                for f in files:
                    counter += 1
                    filestr += str(counter) + ':  ' + f + '<br>'
                    print(str(counter) + ':  ' + f)
                reply('ok')
                app.ChatBot.addAppMsg(filestr)
    # OPEN APPLICATIONS BY VOICE
    elif 'open notepad' in voice_data or 'launch notepad' in voice_data:
        open_application('notepad')

    elif 'open calculator' in voice_data or 'launch calculator' in voice_data:
        open_application('calculator')

    elif 'open word' in voice_data or 'launch word' in voice_data:
        open_application('word')

    elif 'open excel' in voice_data or 'launch excel' in voice_data:
        open_application('excel')

    else:
        reply('I am not functioned to do this!')

    

# ------------------Driver Code--------------------
t1 = Thread(target=app.ChatBot.start)
t1.start()

while not app.ChatBot.started:
    time.sleep(0.5)

wish()
voice_data = None
while True:
    if app.ChatBot.isUserInput():
        voice_data = app.ChatBot.popUserInput()
    else:
        voice_data = record_audio()

    if 'nova' in voice_data:
        try:
            respond(voice_data)
        except SystemExit:
            reply("Exit Successful")
            break
        except:
            print("Exception raised while closing.")
            break
