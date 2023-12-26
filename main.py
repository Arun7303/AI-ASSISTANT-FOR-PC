import os
import webbrowser
import vlc
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import pyjokes
import re
import requests
from newsapi import NewsApiClient
import subprocess
import pyfiglet


newsapi = NewsApiClient(api_key='29790de8a20f4aafa67f7eef8fe0a8e3')

music_player = None

def say(text,rate=150):
    engine = pyttsx3.init()
    engine.setProperty('rate', rate)
    engine.say(text)
    engine.runAndWait()

def take_command():
    r = sr.Recognizer()
    r.pause_threshold = 0.6
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            return "Please try again"
        except sr.RequestError as e:
            return "Sorry, there was an error with the speech recognition service"

def play_music(music_path):
    global music_player
    if music_player is None:
        instance = vlc.Instance()
        player = instance.media_player_new()
        media = instance.media_new(music_path)
        media.get_mrl()
        player.set_media(media)
        music_player = player
        player.play()
    else:
        music_player.stop()
        music_player = None
        say("Music stopped.")

def search_google(query):
    google_url = f"https://www.google.com/search?q={query}"
    webbrowser.open(google_url)
    say(f"Searching Google for {query}")

def search_wikipedia(query):
    say('Searching Wikipedia...')
    results = wikipedia.summary(query, sentences=2)
    say("According to Wikipedia")
    print(results)
    say(results)

def telljoke():
    joke = pyjokes.get_joke()
    say(joke)


def add_numbers(query):
    numbers = re.findall(r'\d+', query)

    if len(numbers) == 2:
        num1 = float(numbers[0])
        num2 = float(numbers[1])
        result = num1 + num2
        response = f"The sum of {num1} and {num2} is {result}"
        say(response)
    else:
        response = "Please provide two numbers to add, for example, 'Add 5 and 3'."
        say(response)

def subtract_numbers(query):
    numbers = re.findall(r'\d+', query)

    if len(numbers) == 2:
        num1 = float(numbers[0])
        num2 = float(numbers[1])
        result = num1 - num2
        response = f"The result of subtracting {num2} from {num1} is {result}"
        say(response)
    else:
        response = "Please provide two numbers to subtract, for example, 'Subtract 10 from 5'."
        say(response)
def multiply_numbers(query):
    numbers = re.findall(r'\d+', query)

    if len(numbers) == 2:
        num1 = float(numbers[0])
        num2 = float(numbers[1])
        result = num1 * num2
        response = f"The result of multiplying {num1} and {num2} is {result}"
        say(response)
    else:
        response = "Please provide two numbers to multiply, for example, 'Multiply 5 by 3'."
        say(response)

def divide_numbers(query):
    numbers = re.findall(r'\d+', query)

    if len(numbers) == 2:
        num1 = float(numbers[0])
        num2 = float(numbers[1])
        if num2 == 0:
            response = "Sorry, division by zero is not allowed."
        else:
            result = num1 / num2
            response = f"The result of dividing {num1} by {num2} is {result}"
        say(response)
    else:
        response = "Please provide two numbers to divide, for example, 'Divide 10 by 2'."
        say(response)

def get_latest_news():
    news_headlines = []
    res = requests.get(
        f"https://newsapi.org/v2/top-headlines?country=in&apiKey=29790de8a20f4aafa67f7eef8fe0a8e3&category=general").json()
    articles = res["articles"]
    for article in articles:
        news_headlines.append(article["title"])
    return news_headlines[:5]

def open_settings():
    subprocess.Popen("C:\\Windows\\System32\\DpiScaling.exe")

def get_random_advice():
    res = requests.get("https://api.adviceslip.com/advice").json()
    return res['slip']['advice']

if __name__ == '__main__':
    print('===========================================================================')
    word=pyfiglet.figlet_format(" Hello...")
    print(word)
    print('===========================================================================')
    say("Hello, I am your AI assistant. How can i help you")
    news_reading = False
    while True:
        print("Listening....")
        query = take_command()

        if "play music" in query:
            music_path = "D:/python/desktop music/music1.mp3"
            play_music(music_path)
        elif "stop music" in query:
            if music_player is not None:
                music_player.stop()
                music_player = None
                say("Music stopped.")
        else:
            sites = [["youtube", "https://www.youtube.com/"], ["wikipedia", "https://www.wikipedia.org/"], 
                    ["google", "https://www.google.com/"], ["erp login", "https://learner.vupune.ac.in/"]
                    , ["new mail", "https://mail.google.com/mail/u/0/#inbox?compose=new"] , ["instagram", "https://www.instagram.com/"]]
            site_opened = False

            for site in sites:
                if f"open {site[0]}" in query:
                    say(f"Opening {site[0]}")
                    webbrowser.open(site[1])
                    site_opened = True

            if "what time is it" in query:
                current_time = datetime.datetime.now().strftime("%H:%M:%S")
                say(f"The current time is {current_time}")
                print(current_time)

            if "open chrome" in query:
                chrome_path=r'"C:\Program Files\Google\Chrome\Application\chrome.exe"'
                say("opening google chrome")
                os.system(f"{chrome_path}")

            if "open word" in query:
                word_path = r'"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"'
                say("Opening Word")
                os.system(f'{word_path}')

            if "open excel" in query:
                excel_path = r'"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"'
                say("Opening Excel")
                os.system(f'{excel_path}')

            if "open powerpoint" in query:
                powerpoint_path = r'"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"'
                say("Opening PowerPoint")
                os.system(f'{powerpoint_path}')

            if re.search(r'\bopen calculator\b', query):
                calc_path = r'C:\Windows\System32\calc.exe'
                say('Opening calculator')
                os.system(f'"{calc_path}"')

            if query.startswith("search google for "):
                search_query = query.replace("search google for ", "")
                search_google(search_query)

            if query.startswith("search wikipedia for "):
                query = query.replace("search wikipedia for ", "")
                search_wikipedia(query)

            if "how r u" in query:
                response = "I am just a computer program, but I'm good. What about you? "
                say(response)

            if "i am not feeling well" in query.lower():
                response = "I'm sorry to hear that. Please take care of yourself and consider visiting a doctor if you don't feel better soon."
                say(response)

            if "i am fine" in query.lower():
                response = "That's great! I'm here to assist you with any tasks or information you need. How can I help you today?"

            if "tell me a joke" in query.lower():
                telljoke()

            if "add" in query:
                add_numbers(query)

            if "subtract" in query:
                subtract_numbers(query)

            if"multiply" in query:
                multiply_numbers(query)

            if "divide" in query:
                divide_numbers(query)

            if "read latest news" in query:
                news_headlines = get_latest_news()
                if news_headlines:
                    say("Here are the latest news headlines:")
                    for index, headline in enumerate(news_headlines, start=1):
                        say(f"{index}. {headline}")
                else:
                    say("Sorry, I couldn't fetch the latest news headlines at the moment.")
            
            if "give me advice" in query:
                answer=get_random_advice()
                say(answer)
            
            if"open settings" in query:
                say("Opening Settings")
                open_settings()

            if "exit assistant" in query.lower() or "close assistant" in query.lower():
                say("Sure.. Exiting Assistant.")
                break 


