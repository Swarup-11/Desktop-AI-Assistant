Desktop AI Assistant 

This repository contains a Python-based AI voice assistant that can perform various tasks, including mathematical calculations, web searches, opening websites, and writing emails. The assistant uses speech recognition and text-to-speech technologies to interact with users.

Features:
Voice Interaction: The assistant listens to your voice commands and responds verbally.
Mathematical Calculations: It can perform basic arithmetic calculations and provide the result.
Open Websites: The assistant can open predefined websites like YouTube, Google, Gmail, and others.
Write Emails: It can help you compose and send emails.
Date, Time, and Day Information: It provides the current date, time, and day on request.
Music Playback: You can command the assistant to play music from a specified file path.

Interact with the Assistant:
The assistant will greet you based on the time of day and ask how it can assist you.
You can give commands such as:
"What is 2 + 2?"
"Open YouTube."
"What is the capital of Maharashtra?"
"What's the time?"
"I want to write an email."

Exit the Assistant:
You can exit the assistant by saying "exit," "quit," or "stop."

Configuration:
Predefined Websites: You can modify the list of predefined websites in the sites list within the main() function.
Capitals Information: The assistant has a predefined list of Indian states and their capitals. This can be modified within the capitals dictionary.
Music File Path: Update the musicpath variable in the main() function to the path where your music file is stored.

Dependencies:
SpeechRecognition: For converting speech to text.
pyttsx3: For text-to-speech conversion.
pypiwin32: For interacting with Windows COM objects.
sympy: For mathematical computations.
webbrowser: To open websites.
subprocess: To start external processes, like playing music.
datetime: To handle date and time operations.

Contribution:
Contributions are welcome! If you have any improvements or suggestions, feel free to create a pull request or open an issue.
