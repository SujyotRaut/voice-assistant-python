import datetime
import json
import random
import re
import subprocess
import webbrowser
from urllib.parse import urlencode

import pyttsx3
import speech_recognition as sr
import wikipedia
from youtube_search import YoutubeSearch

VOICE_ASSISTANT_NAME = 'google'

name = ''

# region Replies
reply_greet = [
    "Hi {name}, what can i do for you",
    "Hello {name}, what can i do for you",
    "What's up {name}, what can i do for you",
]

reply_shutdown = [
    "Bye bye",
    "See yaa",
    "See you",
    "See you later",
]

reply_self_intro = [
    "Hi, I'm an AI voice assistant programed using python",
    f"I'm {VOICE_ASSISTANT_NAME}, A virtual assistant developed for the purpose of cse 3rd year mini project"
]

reply_unknown_cmd = [
    "Sorry i can't understand you",
]
# endregion

# region Commands (Regular Expressions)
cmd_greet = [
    f'(?<=hi {VOICE_ASSISTANT_NAME.lower()}).+',
    f'(?<=hello {VOICE_ASSISTANT_NAME.lower()}).+',
    f'(?<=hey {VOICE_ASSISTANT_NAME.lower()}).+',
    f'(?<=hay {VOICE_ASSISTANT_NAME.lower()}).+',
    f'(?<=ok {VOICE_ASSISTANT_NAME.lower()}).+',
    f'(hi)|(hello)|(hey)|(hay)|(ok) {VOICE_ASSISTANT_NAME.lower()}',
    f'(hi)|(hello)|(hey)|(hay)|(ok)',
    f'(morning)|(afternoon)|(evening)|(night)',
]

cmd_exit = [
    "(exit)|(exit program)|(shutdown)|(bye)|(see you)"
]

cmd_self_intro = [
    r'what are you',
    r'tell me about yourself',
]

cmd_set_username = [
    r'(?<=call me).+',
]

cmd_open_program = [
    r'(?<=open).+',
]

cmd_play_youtube = [
    r'(?<=play).+(?=on youtube)',
    r'(?<=play).+(?=in youtube)',
    r'(?<=play).+',
]

cmd_search_google = [
    r'(?<=search for).+(?=on google)',
    r'(?<=search for).+(?=in google)',
    r'(?<=search for).+',
    r'(?<= search).+(?=on google)',
    r'(?<= search).+(?=in google)',
    r'(?<=search).+',
]

cmd_search_youtube = [
    r'(?<=search for).+(?=on youtube)',
    r'(?<=search for).+(?=in youtube)',
    r'(?<=search).+(?=on youtube)',
    r'(?<=search).+(?=in youtube)',
]

cmd_search_wikipedia = [
    r'(?<=search on wikipedia about).+',
    r'(?<=search in wikipedia about).+',
    r'(?<=search for).+(?=on wikipedia)',
    r'(?<=search for).+(?=in wikipedia)',
    r'(?<=search about).+(?=on wikipedia)',
    r'(?<=search about).+(?=in wikipedia)',
    r'(?<=search).+(?=on wikipedia)',
    r'(?<=search).+(?=in wikipedia)',
    r'(?<=what is).+',
    r'(?<=explain).+',
]
# endregion

# region Text-To-Speech Engine
# Initialize text to speech engine
engine = pyttsx3.init('sapi5')

# Gets you the details of the current voice
voices = engine.getProperty('voices')

# 0-Male voice , 1-female voice
engine.setProperty('voice', voices[1].id)

# Volume between 0 - 1
engine.setProperty('volume', 1)
# endregion


# region Core Functions
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def exit():
    speak(random.choice(reply_shutdown))
    speak('Shutting Down...')


def take_command():
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            print("Listening...")
            recognizer.pause_threshold = 0.5
            audio = recognizer.listen(source)

            print("Recognizing...")
            query = recognizer.recognize_google(audio, language='en-US')

            print("What you said: " + query)
            return str(query).lower()
    except OSError:
        speak("Please give necessary mic permission, and try again")
        return None
    except sr.RequestError:
        speak("It seems your internet connection isn't stable")
        return None
    except sr.UnknownValueError:
        return None


def match_command(regex_list: list, query: str):
    for regex in regex_list:
        match = re.search(regex, query)
        if match != None:
            return match
    return None
# endregion


# region Functions (Features)
def greet(query: str):
    hour = int(datetime.datetime.now().hour)

    if 'night' in query:
        speak(f"Good Night {name}")
        return

    if 'morning' in query and hour >= 0 and hour < 12:
        speak(f"Good Morning {name}, what can i do for you")
        return

    if 'afternoon' in query and hour >= 12 and hour < 18:
        speak(f"Good Afternoon {name}, what can i do for you")
        return

    if 'evening' in query and hour >= 18 and hour < 21:
        speak(f"Good Evening {name}, what can i do for you")
        return

    if 'morning' in query or 'afternoon' in query or 'evening' in query or 'night' in query:
        if hour >= 0 and hour < 12:
            speak(f"I think it's Morning")
        elif hour >= 12 and hour < 18:
            speak(f"I think it's Afternoon")
        else:
            speak(f"I think it's Evening")
        return

    msg = random.choice(reply_greet).replace('{name}', name)
    speak(msg)


def self_intro():
    speak(random.choice(reply_self_intro))


def set_username(query: str):
    global name
    name = match_command(cmd_set_username, query).group().strip()
    speak(f"Ok, I will call you {name} from now on")


def open_program(query: str):
    if 'browser' in query:
        speak("Opening Browser...")
        webbrowser.open('example.com')
    elif 'google' in query:
        speak("Opening Google...")
        webbrowser.open('google.com')
    elif 'youtube' in query:
        speak("Opening Youtube...")
        webbrowser.open('youtube.com')
    elif 'facebook' in query:
        speak("Opening Facebook...")
        webbrowser.open('facebook.com')
    elif 'instagram' in query:
        speak("Opening Instagram...")
        webbrowser.open('instagram.com')
    elif 'notepad' in query:
        speak("Opening Notepad...")
        subprocess.run("notepad")
    elif 'word' in query:
        speak("Opening Microsoft Word...")
        try:
            subprocess.run('winword')
        except:
            speak("Couldn't open microsoft word")
    elif 'excel' in query:
        speak("Opening Microsoft Excel...")
        try:
            subprocess.run('excel')
        except:
            speak("Couldn't open microsoft excel")
    elif 'powerpoint' in query:
        speak("Opening Microsoft Powerpoint...")
        try:
            subprocess.run('powerpnt')
        except:
            speak("Couldn't open microsoft powerpoint")
    else:
        speak("Couldn't open " + re.search(r'(?<=open).+', query).group().strip())


def play_youtube(query: str):
    speak('Opening Youtube....')
    search = match_command(cmd_play_youtube, query).group().strip()
    videos = json.loads(YoutubeSearch(
        search, max_results=10).to_json())['videos']
    if (len(videos) == 0):
        return speak('No videos found for term' + search)
    # video_url = 'https://www.youtube.com' + random.choice(videos)['url_suffix']
    video_url = 'https://www.youtube.com' + videos[0]['url_suffix']
    webbrowser.open(video_url)


def search_google(query: str):
    speak("Searching Google...")
    search = match_command(cmd_search_google, query).group().strip()
    search_base = 'https://www.google.com/search?'
    search_query = urlencode({"q": search})
    search_url = search_base + search_query
    webbrowser.open(search_url)


def search_youtube(query: str):
    speak("Searching Youtube...")
    search = match_command(cmd_search_youtube, query).group().strip()
    search_base = 'https://www.youtube.com/results?'
    search_query = urlencode({"search_query": search})
    search_url = search_base + search_query
    webbrowser.open(search_url)


def search_wikipedia(query: str):
    # speak('Searching Wikipedia...')
    search = match_command(cmd_search_wikipedia, query).group().strip()
    results = wikipedia.summary(search, sentences=5)
    print(results)
    speak("According to Wikipedia")
    speak(results)
# endregion


if __name__ == "__main__":
    while True:
        query = take_command()

        if (query == None):
            continue

        if match_command(cmd_self_intro, query):
            self_intro()
        elif match_command(cmd_search_wikipedia, query):
            search_wikipedia(query)
        elif match_command(cmd_search_youtube, query):
            search_youtube(query)
        elif match_command(cmd_search_google, query):
            search_google(query)
        elif match_command(cmd_open_program, query):
            open_program(query)
        elif match_command(cmd_play_youtube, query):
            play_youtube(query)
        elif match_command(cmd_set_username, query):
            set_username(query)
        elif match_command(cmd_greet, query):
            greet(query)
        elif match_command(cmd_exit, query):
            exit()
            break
        else:
            speak(random.choice(reply_unknown_cmd))
