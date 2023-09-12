import random
import pyttsx3
import speech_recognition as sr
import datetime
import requests
import json
import docx
import os
import webbrowser

name= "PHOENIX. - Personal Helper Optimized for Efficient Navigation and Information eXchange"


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

voice_choice = input('What voice do you want your Assistant to have? ')
if voice_choice.lower() == 'male':
    voice = engine.setProperty('voice', voices[0].id)
else:
    voice = engine.setProperty('voice', voices[1].id)

current_time = datetime.datetime.now().strftime("%H:%M:%S")


salutes = [
    'Im doing great, thanks for asking. How about you?',
    "I'm doing well, thanks. How about yourself?",
    "Not too bad, thanks. How about you?",
    "I'm hanging in there, thanks. How are you doing today?"
]

# sites=[["youtube",'https://www.youtube.com/'],['wikipedia','https://www.wikipedia.org/'],['google','https://www.google.com/']]

def get_weather(city):
    API_key='60cba786e108c30bfd2727f65e4fb91d'
    url = f"https://openweathermap.org/={city}&appid={API_key}&units=metric"
    response = requests.get(url)
    if response.ok:
        try:
            data = json.loads(response.content)
            return data
        except json.decoder.JSONDecodeError as e:
            print(f"Error decoding JSON: {e}")
            return None
    else:
        print(f"Error getting weather data: {response.status_code}")
        return None



text = 'Greetings, Welcome to PHEONIX A.I Version 1.2.2 - Your own AI assistant at your service.'
engine.say(text)
engine.runAndWait()

r = sr.Recognizer()

while True:
    with sr.Microphone() as source:
        r.pause_threshold=1
        print("Speak something...")
        audio = r.listen(source)
    try:
        # Convert speech to text
        text = r.recognize_google(audio)
        print(f"You said: {text}")
    except Exception as e:
        print(f"Error converting speech to text: {e}")

        # Respond to the user
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],['facebook','https://www.facebook.com/'],["Linkedin","https://www.linkedin.com/feed/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in text.lower():
                engine.say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
                
        if "hello" in text.lower():
            engine.say("Hello there!")
        elif "how are you" in text.lower():
            engine.say(random.choice(salutes))
        elif "read file" in text.lower():
            # Ask user for file path
            engine.say("Please provide the file path.")
            engine.runAndWait()
            file_path = input("Enter file path: ")
            # Check if file path is valid
            try:
                doc = docx.Document(file_path)
            except Exception:
                engine.say("Sorry, I couldn't find the file at the given path.")
                engine.runAndWait()
            # Read the contents of the file
            else:
                full_text = ""
                for para in doc.paragraphs:
                    full_text += para.text
                engine.say(f"The contents of the file are: {full_text}")
                engine.runAndWait()

        elif "what time is it" in text.lower():
            engine.say(f"The current time is {current_time}")
        elif 'what is your name' in text.lower():
            engine.say(f'My name is {name}')
        # elif 'open camera'  in text.lower():
        #     open_camera()
        elif "weather" in text.lower():
            # Get the city name from the user
            engine.say("Please specify the city name.")
            engine.runAndWait()
            with sr.Microphone() as source:
                audio = r.listen(source)
            city = r.recognize_google(audio)
            # Get the weather information for the city
            weather_info = get_weather(city)
            engine.say(weather_info)
        elif "remind me" in text.lower():
            # Add reminder functionality here
            pass
        elif f"open excel" in text.lower():
            app_path = os.path.abspath(f"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel")
            os.startfile(app_path)
        elif f"open word" in text.lower():
            app_path = os.path.abspath(f"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Word")
            os.startfile(app_path)
        elif f"open powerpoint" in text.lower():
            app_path = os.path.abspath(f"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\PowerPoint")
            os.startfile(app_path)
        elif "change voice" in text.lower():
            # Change the voice of the assistant dynamically
            new_voice = input("What voice do you want? ")
            if new_voice.lower() == "male":
                voice = engine.setProperty('voice', voices[0].id)
            else:
                voice = engine.setProperty('voice', voices[1].id)
            engine.say("Voice changed successfully.")
        elif 'quit' in text.lower():
            engine.say('Goodbye, It was nice meeting you!')
            engine.runAndWait()
            break
        else:
            print("I can't understand what you said.")
        engine.runAndWait()


        # except sr.UnknownValueError:
        # print("Sorry, I didn't understand what you said.")
        # except sr.RequestError as e:
        # print("Could not request results; {0}".format(e))
 
