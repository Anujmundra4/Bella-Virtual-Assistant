import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import bs4
import requests
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from datetime import date
import newsapi
import pycountry
from bs4 import BeautifulSoup
from win32com.client import Dispatch
import win32com.client as wincl
from urllib.request import urlopen
from newsapi import NewsApiClient
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def password():
    speak("Identify yourself")
    query = takeCommand()
    query = str(query)
    key = "5611"
    if(key==query):
        speak("Access Granted")
    else:
        password()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    assistant_name = ("Bella")
    speak("I am your Assistant")
    speak(assistant_name)


def username():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("Welcome Mr.", uname.center(columns))

    speak("How can i Help you, Sir")




def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    clear()
    password()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()

        if 'send email to' in query:
            try:

                name = list(query.split())

                name = name[name.index('to') + 1]

                speak("what should i say")

                content = takeCommand()

                to = dict[name]

                sendEmail(to, content)

                speak("email has been sent")

            except Exception as e:

                print(e)

                speak("Sorry sir, i was unable to send the mail at the moment. Can you try that again please")


        elif 'wikipedia' in query:

            speak("what do you want to search")
            Query = takeCommand()
            speak("According to wikipedia")
            speak(wikipedia.search(Query))



        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'play music' in query or "play songs" in query:
            speak("opening spotify")
            songs = webbrowser.open("https://open.spotify.com/collection/tracks")

        elif 'open mail' in query:
            speak('working on that sir')
            mail = webbrowser.open("https://mail.google.com/mail/u/0/#inbox")

        elif 'open java' in query:
            speak("Opening Java Compiler")
            os.startfile("C:/Program Files/JetBrains/IntelliJ IDEA Community Edition 2021.2/bin/idea64.exe")

        elif 'open terminal' in query:
            speak("opening command prompt")
            os.system("start cmd")

        elif 'what can you do' in query:
            speak("I can do a lot of things for you sir. I can play music for you, open browsers if you want")
            speak("I can keep you updated with the news, and also, i can send email from your id")


        elif 'what is the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is{strTime}")

        elif 'the date is' in query:
            today = date.today()
            speak(f"Sir, today's date is{today}")


        elif 'how are you' in query:
            speak("Hey, I can't complain. Thanks for asking.")
            speak("By the way, how are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assistant_name = query

        elif "tell me about yourself" in query:
            speak("Hii, I am Bella. I was created as a part of project known as virtual assistant and i am really thankful for that")
            speak("But now, I can do many things to make things easy for my owner")


        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif 'search' in query or 'play' in query:

            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.open(query)


        elif "open notepad" in query:
            speak("opening notepad")
            osCommandString = "notepad.exe"
            os.system(osCommandString)

        elif "open powerpoint" in query:
            speak("opening power point presentation")

            os.startfile('C:/Users/91772/OneDrive/Desktop/PowerPoint.pptx')


        elif 'world news' in query:

            def NewsFromBBC():

                query_params = {
                    "source": "bbc-news",
                    "sortBy": "top",
                    "apiKey": "8bbccf9e61ab45bc803051260f090e69"
                }
                main_url = " https://newsapi.org/v1/articles"

                res = requests.get(main_url, params=query_params)
                open_bbc_page = res.json()

                article = open_bbc_page["articles"]

                results = []

                for ar in article:
                    results.append(ar["title"])

                for i in range(len(results)):

                    print(i + 1, results[i])

                speak = Dispatch("SAPI.Spvoice")
                speak.Speak(results)


        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")


        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")


        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])


        elif "Bella" in query:

            wishMe()
            speak("Bella in you service sir.")
            speak(assistant_name)



        elif 'weather' in query:


            api_key = "API KEY"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"

            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid=" + '15dda5fe83de140357727be28f252c60' + "&q=" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["cod"] != "404":
                y = x["main"]

                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]

                weather_description = z[0]["description"]

                speak(" Temperature (in kelvin unit) = " + str(
                current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                current_pressure) + "\n humidity (in percentage) = " + str(
                current_humidiy) + "\n description = " + str(weather_description))

            else:
                speak(" City Not Found ")


        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(assistant_name)
