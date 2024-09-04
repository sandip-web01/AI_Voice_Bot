'''from flask import Flask, request, render_template,jsonify
import speech_recognition as sr
import datetime
import webbrowser
import os
import pythoncom
import win32com.client as win32

app = Flask(__name__)

# Initialize the recognizer
recognizer = sr.Recognizer()

#Home page of web application
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run-python', methods=['POST'])
def run_python():
    wish_me()
    while True:
        query = listen()
        if query == "none":
            continue
        if "exit" in query or "stop" in query:
            speak("Thank you. You are a good Speaker. Good bye ! Have a nice time.")
            break
        perform_task(query)
    return render_template('index.html')

def listen():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

        try:
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language='en-US')
            print(f"User said: {query}")
            return query.lower()
        
        except sr.UnknownValueError:
            speak("Sorry, I didnot catch that. Could you please repeat?")
            return "none"
        
        except sr.RequestError:
            speak("Sorry, the service is down. Please try again later.")
            return "none"
        
        except Exception as e:
            speak("An unexpected error occurred.")
            return "none"
        
def wish_me():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am your personal assistant. How can I help you today?")

def create_word_document(data):
    # Initialize COM automation
    pythoncom.CoInitialize()

    # Create a new instance of Word application
    word_app = win32.Dispatch("Word.Application")
    word_app.Visible = True

    # Add a new document
    doc = word_app.Documents.Add()

    # Write data to the document
    doc.Content.Text = data

def wr_data():
    query =listen()
    return query

def write_data(app_name):
    # Listen to what the user wants to write
    query = listen()
    
    if app_name == "notepad":
        with open("output.txt", "w") as file:
            file.write(query)
        os.system("notepad output.txt")
    elif app_name in ["word", "excel", "powerpoint"]:
        # For Word, Excel, and PowerPoint, writing directly isn't trivial using voice commands,
        # as it requires manipulating these apps' APIs or automation tools.
        speak("Currently, I cannot automatically write to Word, Excel, or PowerPoint. Please enter the data manually.")
    else:
        speak("Sorry, I can't write to this application right now.")

def open_website(query):
    if "open" in query:
        query = query.replace("open", "")
        website = query.strip().replace(" ", "")
        url = f"https://{website}.com"
        speak(f"Opening {website}")
        webbrowser.open(url)

def search_google(query):
    query = query.replace(" ", "+")
    url = f"https://www.google.com/search?q={query}"
    webbrowser.open(url)

def play_song_on_youtube(song_name):
    query = song_name.replace(" ", "+")
    url = f"https://www.youtube.com/results?search_query={query}"
    webbrowser.open(url)

def perform_task(query):
    if "time" in query:
        str_time = datetime.datetime.now().strftime("%I:%M %p")
        print(f"Time: {str_time}")
        speak(f"The time is {str_time}")
    # Open and write into Notepad
    elif "open notepad" in query:
        speak("Opening Notepad. What would you like to write?")
        os.system("notepad.exe")
        write_data("notepad")
    # Open Word and write data
    elif "open word" in query:
        speak("Opening Microsoft Word. What would you like to write?") 
        # Replace with dynamic input or response logic
        create_word_document(wr_data())

    # Open Excel and write data
    elif "open excel" in query:
        speak("Opening Microsoft Excel. What data would you like to input?")
        os.system("start excel")
        write_data("excel")

    # Open PowerPoint and write data
    elif "open powerpoint" in query:
        speak("Opening Microsoft PowerPoint. What would you like to include in your presentation?")
        os.system("start powerpnt")
        write_data("powerpoint")

    # Open VS Code
    elif "open vscode" in query or "open visual studio code" in query:
        speak("Opening Visual Studio Code.")
        os.system("code")
    elif "open whatsapp" in query:
        speak("Opening WhatsApp")
        os.system("Whatsapp.exe")
    elif "search for" in query:
        search_term = query.replace("search for", "").strip()
        speak(f"Searching for {search_term} on Google")
        search_google(search_term)
    elif "play" in query and "youtube" in query:
        song_name = query.replace("play", "").replace("on youtube", "").strip()
        speak(f"Playing {song_name} on YouTube")
        play_song_on_youtube(song_name)
    elif "open" in query:
        website = query.replace("open", "").strip().replace(" ", "")
        speak(f"Opening {website}")
        url = f"https://{website}.com"
        webbrowser.open(url)

    elif "say" in query:
        query = query.replace("say", "").strip()
        speak(query)

    else:
        speak("Sorry, I can't help with that right now.")

def speak(text):
    os.system(f'powershell -c "Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak(\'{text}\');"')

if __name__ == "__main__":
    app.run(debug=True)
'''
from flask import Flask, request, render_template
import speech_recognition as sr
import time
import datetime
import webbrowser
import os
import pythoncom
import win32com.client as win32
from openpyxl import Workbook
'''from config import get_database'''

app = Flask(__name__)

# Initialize the recognizer
recognizer = sr.Recognizer()

'''db = get_database()'''

# Home page of web application
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run-python', methods=['POST'])
def run_python():
    wish_me()
    run_once = 1
    query = ""
    while run_once == 1:
        run_once += 1
        query = listen()
        if query in ["exit", "stop"]:
            speak("Thank you. You are a good speaker. Goodbye! Have a nice time.")
            break
        perform_task(query)
    return render_template('index.html', command=query)

def listen(slowdown_factor=1):
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

        try:
            time.sleep(slowdown_factor)
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language='en-US')
            print(f"User said: {query}")
            return query.lower()
        
        except sr.UnknownValueError:
            speak("Sorry, I did not catch that. Could you please repeat?")
        except sr.RequestError:
            speak("Sorry, the service is down. Please try again later.")
        return ""

def wish_me():
    hour = datetime.datetime.now().hour
    greeting = "Good Morning!" if hour < 12 else "Good Afternoon!" if hour < 18 else "Good Evening!"
    speak(greeting)
    speak("I am your personal assistant. How can I help you today?")

def perform_task(query):
    if "time" in query:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"The time is {current_time}")
    #Open Notepad and write what the user says and update it in text document.
    elif "open notepad" in query:
        speak("Opening Notepad. Start dictating, and I will write for you. Say stop writing when you are done.")
        write_to_notepad()
    #Open word and write what the user says and update it in word document.
    elif "open word" in query:
        speak("Opening Microsoft Word. Start dictating, and I will write for you. Say stop writing when you are done.")
        create_word_document()

    elif "open excel" in query:
        speak("Opening Microsoft Excel.")
        write_to_excel()
    elif "open powerpoint" in query:
        speak("Opening Microsoft PowerPoint.")
        os.system("start powerpnt")
    elif "open vscode" in query or "open visual studio code" in query:
        speak("Opening Visual Studio Code.")
        os.system("code")
    elif "open whatsapp" in query:
        speak("Opening WhatsApp.")
        os.system("WhatsApp")
    elif "search for" in query:
        search_term = query.replace("search for", "").strip()
        speak(f"Searching for {search_term} on Google")
        webbrowser.open(f"https://www.google.com/search?q={search_term.replace(' ', '+')}")
    elif "play" in query and "youtube" in query:
        song_name = query.replace("play", "").replace("on youtube", "").strip()
        speak(f"Playing {song_name} on YouTube")
        webbrowser.open(f"https://www.youtube.com/results?search_query={song_name.replace(' ', '+')}")
    elif "open" in query:
        website = query.replace("open", "").strip().replace(" ", "")
        speak(f"Opening {website}")
        webbrowser.open(f"https://{website}.com")
    elif "say" in query:
        speak(query.replace("say", "").strip())
    else:
        speak("Sorry, I can not help with that right now.")

def speak(text):
    os.system(f'''powershell -c "Add-Type -AssemblyName System.Speech; $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer;$speak.SelectVoice('Microsoft Zira Desktop');  $speak.Speak(\'{text}\');"''')

def write_to_notepad():
    # Open Notepad
    os.system("notepad.exe")  # Wait for Notepad to open

    while True:
        query = listen()
        if "stop writing" in query:
            speak("Stopped writing.")
            break
        else:
            # Send the recognized text to Notepad
            os.system(f'''powershell -c "$wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('{query} ')"''')

def create_word_document():
    pythoncom.CoInitialize()
    word_app = win32.Dispatch("Word.Application")
    word_app.Visible = True
    doc = word_app.Documents.Add()
    paragraph = doc.Content.Paragraphs.Add()
    
    while True:
        query = listen()
        if "stop writing" in query:
            speak("Stopped writing.")
            break
        else:
            paragraph.Range.Text += query + " "
            doc.Content.InsertAfter(query + " ")
            paragraph.Range.Collapse(0)  # Move cursor to end of document
            word_app.Application.ScreenUpdating = True

def write_to_excel():
    # Create a new workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "User Data"

    # Ask for the data to write in Excel
    speak("What data would you like to enter into Excel?")
    data = listen()

    # Writing each sentence or data item into a separate row
    rows = data.split(",")  # Assuming user inputs comma-separated data for multiple rows
    for index, row in enumerate(rows, start=1):
        # Split each row into columns (assuming space-separated values)
        columns = row.strip().split(" ")
        for col_index, value in enumerate(columns, start=1):
            ws.cell(row=index, column=col_index).value = value

    # Save the workbook
    filename = "user_data.xlsx"
    wb.save(filename)
    speak(f"Data has been saved in {filename}")

    # Open the Excel file automatically
    os.system(f'start excel "{filename}"')
def write_data(app_name):
    data = listen()
    if app_name == "notepad":
        with open("output.txt", "w") as file:
            file.write(data)
        os.system("notepad output.txt")
    else:
        speak("Currently, writing to Word, Excel, or PowerPoint is not automated.")

if __name__ == "__main__":
    app.run(debug=True)