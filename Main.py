import speech_recognition as sr
import os
import pyttsx3
import win32com.client
import webbrowser
import subprocess #For music start
import datetime
import threading
import sympy as sp  #Used for mathimatical canculations.

def speak(text):
    speaker=win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def take_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.pause_threshold = 0.6
        print("Listening...")
        audio=recognizer.listen(source)
        try:
            print("Recognizing....")
            query=recognizer.recognize_google(audio, language="en-in")
            print(f"User said:{query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            speak("Sorry, I did not understand that. Please try again.")
            return None
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            speak("Could not request results. Please check your network connection and try again.")
            return None
        
def get_input(prompt):  #Used for input speaking.
    speak(prompt)
    return input(prompt)

def write_email():
    speak("Ready to write email sir...")
    recipient = get_input("Who would you like to send the email to?")
    subject = get_input("What’s the subject of the email?")
    body = get_input("What do you want to say in the email?")

    email = f"""To: {recipient}. 
              Subject: {subject}
              
              {body}"""
    speak("Here's yout email.")
    speak(email)

    confirm = get_input("Would you like to send this email? (y/n)")
    if confirm.lower() == "y":
        speak("Your email has been sent!!")
    else:
        speak("Your email has not been sent!!")
    
def math_calculations(expression):
    try:
        expression = expression.lower().replace("what is", "").strip()
        result = sp.sympify(expression).evalf()

        if result == int(result):
            result = int(result)
        else:
            result = round(float(result), 2)

        print(f"Calculated result: {result}")
        speak(f"The result is {result}")
        return str(result)
    except Exception as e:
        print(f"Error: {e}")
        speak("Sorry, I couldn't calculate that. Please check the expression.")


def main():
    current_hour = datetime.datetime.now().hour

    if 5<=current_hour<12:
        speak("Good morning swarup")
    elif 12<=current_hour<17:
        speak("Good afternoon swarup")
    elif 17<=current_hour<22:
        speak("Good evening swarup")
    else:
        speak("Good night swarup")
    speak("Tell me how can I assist you today?")
    while True:
        query=take_command()
        if query:
            if "exit" in query or "quit" in query or "stop" in query:
                speak("Take care and enjoy the rest of your day!")
                break

        sites=[["youtube","https://www.youtube.com"],
               ["wikipedia","https://www.wikipedia.com"],
               ["google","https://www.google.com"],
               ["gmail","https://mail.google.com/mail/u/0/?pli=1#inbox"],
               ["opportunity","https://pod.ai/"],
               ["result","https://exam.mitapps.in/landing"],
               ["instagram","https://www.instagram.com/"],
               ["amazon","https://www.amazon.in/"],
               ["netflix","https://www.netflix.com/in/"],
               ["live score","https://crex.live/"]]
        opened = False
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
                opened = True
        
        capitals = {"Maharashtra":"Mumbai","Goa":"Panjim","Andhra Pradesh":"Amaravati",
                    "Arunachal Pradesh":"Itanagar","Assam":"Dispur","Bihar":"Patna",
                    "Chhattisgarh":"Raipur","Gujarat":"Gandhinagar","Haryana":"Chandigarh",
                    "himachal Pradesh":"Shimla","Jharkhand":"Ranchi","Karnataka":"Bangalore",
                    "Kerala":"Thiruvananthapuram","Madhya Pradesh":"Bhopal","Manipur":"Imphal",
                    "Meghalaya":"Shilong","Mizoram":"Aizawl","Nagaland":"Kohima",
                    "Odisha":"Bbhubaneswar","Punjab":"Chandigarh","Rajasthan":"Jaipur",
                    "Sikkim":"Gangtok","Tamil Nadu":"Chennai","Telangana":"Hydrabad",
                    "Tripura":"Agartala","Uttarakhand":"Dehradun","Uttar Pradesh":"Lucknow",
                    "West Bengal":"Kolkata"}
        opened = False
        for state,capital in capitals.items():
            if f"the capital of {state}".lower() in query.lower():
                speak(f"The capital of {state} is {capital}")
                opened = True


        if "open music" in query:
            musicpath = "C:\\Users\\Admin\\OneDrive\\Music\\Satranga(PagalWorld.com.cm).mp3"
            speak(f"Playing music sir...")
            os.startfile(musicpath)
            opened = True

        if "the time" in query:
            musicpath = "C:\\Users\\Admin\\OneDrive\\Music\\Satranga(PagalWorld.com.cm).mp3"
            str_time = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"Sir it's {str_time}")
            opened = True

        if "the date" in query:
            str_date = datetime.datetime.now().strftime("%Y-%m-%d")
            speak(f"Sir, today's date is {str_date}")
            opened = True
        
        if "the day" in query:
            str_day = datetime.datetime.now().strftime("%A")
            speak(f"Sir, today's day is {str_day}")
            opened = True

        if "open vs code" in query:
            vspath = "C:\\Users\\Admin\\OneDrive\\Desktop\\Visual Studio Code.lnk"
            speak(f"Opening visual studio code sir...")
            os.startfile(vspath)
            opened = True

        if "search for" in query:
            search_query = query.split("search for")[1].strip()
            webbrowser.open(f"https://www.google.com/search?q={search_query}")
            speak(f"Searching for {search_query}")
            opened = True

        if "want to write email" in query:
            write_email()
            opened = True

        '''if any(char.isdigit() for char in query):
            result = math_calculations(query)
            speak(f"The result is {result}")
            opened  = True'''

        if any(char.isdigit() for char in query):
            math_calculations(query)
            opened = True

        if not opened:
            speak(f"You said: {query}")
            speak("Hmmm, that doesn't seem familiar. Maybe try asking in a different way?")
            
        
        

if __name__== '__main__':
    main()


#pip install speechreconginition
#pip install wikipedia
#pip install openai