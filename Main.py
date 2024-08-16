import speech_recognition as sr
import os
import pyttsx3
import win32com.client
import webbrowser
import subprocess #For music start
import datetime
import threading

'''def speak(text):
    speaker=win32com.client.Dispatch("SAPI.SpVoice")
    speaker.speak(text)'''

def speak(Text):
    engine = pyttsx3.init("sapi5")
    voice = engine.getProperty("voices")
    engine.setProperty("voice",voice[0].id)
    engine.setProperty("rate",200)
    print(f"jarvis: {Text}")
    engine.say(text=Text)
    engine.runAndWait()

def take_command():
    recognizer=sr.Recognizer()
    with sr.Microphone() as source:
        #recognizer.pause_threshold=0.8
        audio=recognizer.listen(source)
        try:
            print("Recognizing....")
            query=recognizer.recognize_google(audio, language="en-in")
            print(f"User said:{query}")
            return query
        except Exception as e:
            return "Some error occured. Sorry from Jarvis"
        

if __name__== '__main__':
    current_hour = datetime.datetime.now().hour

    if 5<=current_hour<12:
        speak("Good morning swarup")
    elif 12<=current_hour<17:
        speak("Good afternoon swarup")
    elif 17<=current_hour<24:
        speak("Good evening swarup")
    speak(f"It's {datetime.datetime.now().strftime("%I:%M %p")}")
    speak("Tell me how can I assist you today?")
    while True:
        print("Listening...")
        query=take_command()

        sites=[["youtube","https://www.youtube.com"],["wikipedia","https://www.wikipedia.com"],
               ["google","https://www.google.com"],["gmail","https://mail.google.com/mail/u/0/?pli=1#inbox"],
               ["opportunity","https://pod.ai/"],["result","https://exam.mitapps.in/landing"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
            
        '''cities = [["Pune"],["Mumbai"],["Nagpur"],["Thane"],["Nashik"],["Solapur"],
                  ["Kolhapur"],["Satara"],["Nanded"]]
        for city in cities:
            if f"Open {city[0]}".lower() in query.lower():
                speak(f"Today's whether in {city[0]} is ")'''
        
        capitals = {"Maharashtra":"Mumbai","Goa":"Panjim"}
        for state,capital in capitals.items:
            if f"the capital of {state[0]}".lower() in query.lower():
                speak(f"The capital of {state[0]} is {capital[0]}")


        if "open music" in query:
            musicpath = "C:\\Users\\Admin\\OneDrive\\Music\\Satranga(PagalWorld.com.cm).mp3"
            speak(f"Opening music sir...")
            os.startfile(musicpath)
            
        elif "the time" in query:
            musicpath = "C:\\Users\\Admin\\OneDrive\\Music\\Satranga(PagalWorld.com.cm).mp3"
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            speak(f"Sir it's {hour} {min}")

        elif "open vs code" in query.lower():
            vspath = "C:\\Users\\Admin\\OneDrive\\Desktop\\Visual Studio Code.lnk"
            speak(f"Opening visual studio code sir...")
            os.startfile(vspath)

        elif "hello jarvis, how are you?" in query:
            speak("I'm fine. What about you?")
        else:
            exit()


#pip install speechreconginition
#pip install wikipedia
#pip install openai