import speech_recognition as sr
import pyaudio
import pyttsx3
import datetime
import wikipedia
import webbrowser
import win32com.client as win32
import os
import smtplib
import argparse
import requests
from bs4 import BeautifulSoup as bs

USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36"
LANGUAGE = "en-US,en;q=0.5"

def get_weather_data(url):
    session = requests.Session()
    session.headers['User-Agent'] = USER_AGENT
    session.headers['Accept-Language'] = LANGUAGE
    session.headers['Content-Language'] = LANGUAGE
    html = session.get(url)
    soup = bs(html.text, "html.parser")

    result = {}
    result['region'] = soup.find("div", attrs={"id": "wob_loc"}).text
    result['temp_now'] = soup.find("span", attrs={"id": "wob_tm"}).text
    result['dayhour'] = soup.find("div", attrs={"id": "wob_dts"}).text
    result['weather_now'] = soup.find("span", attrs={"id": "wob_dc"}).text
    result['precipitation'] = soup.find("span", attrs={"id": "wob_pp"}).text
    result['humidity'] = soup.find("span", attrs={"id": "wob_hm"}).text
    result['wind'] = soup.find("span", attrs={"id": "wob_ws"}).text 

    return result

parser = argparse.ArgumentParser(description="Quick Script for Extracting Weather data using Google Weather")
parser.add_argument("region", nargs="?", help="""Region to get weather for, must be available region.
                                    Default is your current location determined by your IP Address""", default="")

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)
engine.setProperty('rate', 160)
engine.setProperty('volume', 1.0)


def speak(text):
    engine.say(text)
    engine.runAndWait()

def startup_statement():
    hour = int(datetime.datetime.now().hour)
    
    if(hour > 0 and hour <= 12):
        speak("Good morning")
    else:
        speak("Good afternoon")

    speak("What can my services provide?")

def waiting_command():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        try:
            r.adjust_for_ambient_noise(source, duration=5)
            print("Waiting...")
            audio = r.listen(source, timeout=5)
            speaker = r.recognize_google(audio)
            print(speaker)

        except:
            print("Timeout, running again")
            speaker = ""

        return speaker

def speaker_input():
    r = sr.Recognizer()

    with sr.Microphone() as source:
       print("Listening...")
       audio = r.listen(source)

    try:
        speaker = r.recognize_google(audio)

    except:
        speak("Translation Failed")
        speaker = ""

    return speaker

def SendEmail(emailAddress, subject, message):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = emailAddress
    mail.Subject = subject
    message = message.encode("raw_unicode_escape")
    message = message.decode("unicode_escape")
    mail.Body = message
    # Attach a file    
    mail.Send()

quit_commands = ["quit", "stop", "terminate", "kill yourself", "end", "shut down"]
search_commands = ["search", "google", "who is", "look up", "wikipedia", "tell me about"]
music_commands = ["music", "play", "song", "sing to me"]
weather_commands = ["weather"]
email_commands = ["email"]
open_commands = ["open", "start"]


while (1):
    running_in_back = waiting_command()

    k = 0
    while k < len(quit_commands):
        if quit_commands[k] in running_in_back:
            exit(0)
        k += 1       

    if 'Saturn' in running_in_back:
        startup_statement()
        question = speaker_input()

        print(question)

        k = 0
        while k < len(quit_commands):
            if quit_commands[k] in running_in_back:
                exit(0)
            k += 1

        k = 0
        while k < len(search_commands):
            if search_commands[k] in question:
                speak("Searching")
                question = question.replace(search_commands[k], "")
                results = wikipedia.summary(question, sentences =2)
                speak(results)
            k += 1

        k = 0
        while k < len(music_commands):
            if music_commands[k] in question:
                artist_dir = "C:\\Users\\Nkitc\\Music\\iTunes\\iTunes Media\\Music"
                artist_list = os.listdir(artist_dir)
                print(artist_list)
                speak("Do you have a perfered song, album, or artist?")
                answer = speaker_input()

                if 'song' in answer:
                    speak("Please say the name of the song.")
                    song = speaker_input()
                elif 'artist' in answer:
                    speak("Please say the name of the artist.")
                    artist = speaker_input()
                    i = 0
                    while i < len(artist_list):

                        if artist_list[i] in artist:
                            playsong = artist_dir + "\\" + artist
                            print(playsong)
                            startsong = os.listdir(playsong)
                            print(startsong)

                            count = len(startsong)

                            if len(startsong) == 1:
                                playsong += "\\" + startsong[0]
                                startsong = os.listdir(playsong)
                                os.startfile(os.path.join(playsong, startsong[0]))
                            else:
                                speak("There are " + str(count) +" albums")
                                speak(startsong)
                                speak("Which one would you like to listen too")
                                album_choice = speaker_input()
                                print(album_choice)

                                spoken_numbers = ["one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"
                                "first", "second", "third", "forth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth"]
                        
                                position = 0
                                while position <= count:
                        
                                    if spoken_numbers[position] in album_choice:
                                        actual_song = playsong + "\\" + startsong[position]
                                        actual_start_song = os.listdir(actual_song)
                                        os.startfile(os.path.join(actual_song, actual_start_song[0]))

                                    position += 1
                        i += 1
                               
                elif 'album' in answer:
                    speak("Please say the name of the album.")

            k += 1

        k = 0
        while k < len(weather_commands):
            if weather_commands[k] in question:

                URL = "https://www.google.com/search?lr=lang_en&ie=UTF-8&q=weather"
                
                speak("Please say the city where you want the weather from")
                
                location = speaker_input()

                URL += "-" + location

                print(URL)
                
                data = get_weather_data(URL)

                speak("Weather in:" + data["region"])
                speak("Today is" + data["dayhour"])
                speak(f"The temperature is {data['temp_now']}" + "degrees")
                speak("The skys are:" + data['weather_now'])
                speak("Humidity is at" + data["humidity"])
                speak("Wind speeds up too" + data["wind"][0] + "miles per hour")
                print("Weather in:" + data["region"])
                print("Today is" + data["dayhour"])
                print(f"The temperature is {data['temp_now']}" + "degrees")
                print("The skys are:" + data['weather_now'])
                print("Humidity is at" + data["humidity"])
                print("Wind speeds up too " + data["wind"][0] + " miles per hour")#issues saying the winds speed

            k += 1



        k = 0
        while k < len(email_commands):
            if email_commands[k] in question:
                speak("sending email")
                SendEmail("rcoote@microsoft.com", "Adobe Creative Cloud", "Let me know when you want to download Adobe on your machine.")

            k += 1

        k = 0
        while k < len(open_commands):
            if open_commands[k] in question:
                speak("what would you like to open?")
                responce = speaker_input()

                if 'Adobe' in responce:
                    speak("Opening photoshop")
                    os.startfile("C:\Program Files\Adobe\Adobe Photoshop 2020/Photoshop.exe")

            k += 1


#to install using pip the command is:
#python -m pip install SpeechRecognition
#python -m pip install pipwin
#python -m pipwin install pyaudio
#pip install pyttsx3==2.6
#pip install wikipedia
#We need these libraries to use the speech libraries
