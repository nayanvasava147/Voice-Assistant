from flask import Flask, render_template, jsonify 
import os
import webbrowser
import threading
import datetime
import pyttsx3
import speech_recognition as sr
import subprocess
import glob
import cv2
import time
import requests
import urllib.parse
import pyautogui
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from bs4 import BeautifulSoup
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import GUID  

IID_IAudioEndpointVolume = GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}")

app = Flask(__name__)

class VoiceAssistant:
    def __init__(self):
        self.engine = pyttsx3.init()
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[1].id)
        self.silent_mode = False
        self.output_dir = "recordings"
        os.makedirs(self.output_dir, exist_ok=True)

        self.app_keywords = {
            "instagram": "https://www.instagram.com",
            "snapchat": "https://www.snapchat.com",
            "whatsapp": "https://web.whatsapp.com",
            "netflix": "https://www.netflix.com/browse",
            "anime": "https://hianime.to/home",
            "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "edge": "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
            "excel": "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Excel.lnk"
        }

        self.app_names = self.get_installed_apps()

    
    def get_installed_apps(self):
        program_files = os.path.join(os.getenv('ProgramData'), 'Microsoft', 'Windows', 'Start Menu', 'Programs')
        return [os.path.splitext(os.path.basename(f))[0].lower() for f in glob.glob(os.path.join(program_files, '**', '*.lnk'), recursive=True)]

    def speak(self, audio):
        if not self.silent_mode:
            self.engine.say(audio)
            self.engine.runAndWait()

    def take_command(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            try:
                audio = r.listen(source, timeout=4, phrase_time_limit=3)
            except sr.WaitTimeoutError:
                self.speak("I didn't hear anything.")
                return "none"
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            self.speak("Sorry, I didn't catch that.")
            return "none"
        except sr.RequestError:
            self.speak("Couldn't connect to the recognition service.")
            return "none"

    def handle_silent_mode(self, query):
        silent_commands = ["chup raho", "shut up", "quiet", "stay silent", "silent"]
        wake_commands = ["system", "hello system", "hello assistant"]

        if any(word in query for word in silent_commands):
            self.speak("Sorry to interrupt, going silent.")
            self.silent_mode = True
            return True
        elif any(word in query for word in wake_commands):
            self.silent_mode = False
            self.speak("Yeah, I'm here. How can I help you?")
            time.sleep(2)
            return False
        return False

    def change_volume(self, action, value=None):
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            current_volume = volume.GetMasterVolumeLevelScalar() * 100
            self.speak(f"Current volume is {int(current_volume)}%")

            if action == "increase":
                target_volume = min(100, current_volume + (value if value else 10))
            elif action == "decrease":
                target_volume = max(0, current_volume - (value if value else 10))
            elif value is not None:
                target_volume = min(100, value)

            volume.SetMasterVolumeLevelScalar(target_volume / 100, None)
            self.speak(f"Volume set to {int(target_volume)}%")

        except Exception as e:
            self.speak(f"An error occurred while adjusting the volume: {str(e)}")

    # Function to open camera and handle click and delete operations
    def camera(self, query):
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            self.speak("Could not open the camera.")
        else:
            picture_taken = False
            img_path = None

        while True:
            ret, img = cap.read()
            if not ret:
                self.speak("Failed to grab frame.")
                break
            
            cv2.imshow('Camera Feed', img)

            if "click picture" in query and not picture_taken:
                img_path = "captured_image.jpg"
                cv2.imwrite(img_path, img)
                self.speak("Picture clicked!")
                picture_taken = True

            if picture_taken:
                # Display the captured image
                img_display = cv2.imread(img_path)
                cv2.imshow('Captured Image', img_display)

                if "delete" in query:
                    if os.path.exists(img_path):
                        os.remove(img_path)
                        self.speak("Picture deleted.")
                    else:
                        self.speak("No picture to delete.")
                    picture_taken = False
                    img_path = None
                    cv2.destroyWindow('Captured Image')  # Close the captured image window

            # Exit when 'Esc' key is pressed
            if cv2.waitKey(1) & 0xFF == 27:
                break
        cap.release()
        cv2.destroyAllWindows()

    def take_screenshot(self):
        screenshot = pyautogui.screenshot()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filepath = os.path.join(self.output_dir, f"screenshot_{timestamp}.png")
        screenshot.save(filepath)
        self.speak(f"Screenshot saved as {filepath}")

    def open_app_or_website(self, query):
        for app_name, path in self.app_keywords.items():
            if app_name in query:
                if "close" in query:
                    os.system(f"taskkill /f /im {app_name}.exe")
                    self.silent_mode = False
                    self.speak(f"{app_name.capitalize()} closed.")
                else:
                    if path.startswith("http"):
                        webbrowser.open(path)
                    else:
                        os.startfile(path)
                    self.silent_mode = True
                    self.speak(f"{app_name.capitalize()} opened.")
                return

        for app_name in self.app_names:
            if app_name in query:
                try:
                    app_path = os.path.join(os.getenv('ProgramData'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', app_name + '.lnk')
                    subprocess.Popen(app_path)
                    self.silent_mode = True
                    self.speak(f"{app_name.capitalize()} opened.")
                    return
                except Exception as e:
                    self.speak(f"Failed to open {app_name}.")

        site_name = query.replace("open ", "").strip()
        default_url = f"http://www.{site_name}.com"
        self.speak(f"Opening {site_name}.com.")
        webbrowser.open(default_url)

    def youtube_search(self):
        webbrowser.open("https://www.youtube.com")
        self.speak("YouTube is open. What would you like to search?")
        for attempt in range(3):
            search_query = self.take_command()
            if search_query and search_query != "none":
                youtube_search_url = f"https://www.youtube.com/results?search_query={search_query.replace(' ', '+')}"
                self.speak(f"Searching YouTube for {search_query}.")
                webbrowser.open(youtube_search_url)
                return
            elif attempt < 2:
                self.speak("I didn't catch that. Please tell me what you'd like to search on YouTube.")
        self.speak("I couldn't hear your search query. Please try again later.")

    def wikipedia_search(self, query):
        topic = query.replace("search", "").strip()
        wikipedia_url = f"https://en.wikipedia.org/wiki/{topic.replace(' ', '_')}"
        self.speak(f"Searching Wikipedia for {topic}.")
        webbrowser.open(wikipedia_url)

    def search_bing(self, query):
        encoded_query = urllib.parse.quote(query)
        bing_url = f"https://www.bing.com/search?q={encoded_query}"
        webbrowser.open(bing_url)
        self.speak(f"Searching Bing for {query}")

    def google_search(self, query):
        api_key = "AIzaSyDWqFsaV3PS5oW6HkKEcgjij4eklTtxOVk"
        search_engine_id = "YOUR_SEARCH_ENGINE_ID"  # Replace with your actual Search Engine ID
        url = f"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={search_engine_id}&q={query}"
        
        response = requests.get(url)
        if response.status_code == 200:
            results = response.json()
            if 'items' in results:
                for item in results['items'][:3]:  # Limit to the first 3 items for brevity
                    title = item.get("title")
                    snippet = item.get("snippet")
                    link = item.get("link")
                    self.speak(f"Title: {title}. {snippet}. You can find it at {link}.")
            else:
                self.speak("No results found for your search.")
        else:
            self.speak("I encountered an error while searching. Please try again later.")

    def anime(self, query):
        topic = query.replace("anime", "").strip()
        if topic:
            search_url = f"https://hianime.to/search?keyword={topic}"
            self.speak(f"Searching HiAnime for {topic}.")
            webbrowser.open(search_url)
        else:
            self.speak("Please specify the anime title you want to search for.")

    def execute_query(self, query):
        if self.handle_silent_mode(query):
            return
        if "anime" in query:
            self.anime(query)
        elif "youtube" in query:
            self.youtube_search()
        elif "search" in query:
            self.wikipedia_search(query)
        elif "google" in query:
            self.google_search(query)
        elif "open" in query:
            self.open_app_or_website(query)
        elif "search bing" in query:
            self.search_bing(query)
        elif "screenshot" in query:
            self.take_screenshot()
        elif "volume" in query:
            self.change_volume("current")
        elif "increase volume" in query:
            self.change_volume("increase")
        elif "decrease volume" in query:
            self.change_volume("decrease")
    
        if "stop" in query or "exit" in query:    
            self.speak("Shutting down. Have a great day!")
            exit()

        if "close" in query:
            self.speak("Please specify which app to close.")

         

    def run(self):
        self.speak("   Starting voice assistant. How can I assist you?")
        while True:
            query = self.take_command()
            if query != "none":
                self.execute_query(query)
        

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/start_assistant', methods=['POST'])
def start_assistant():
    # Start the assistant in a separate thread
    def run_assistant():
        assistant = VoiceAssistant()
        assistant.run()

    assistant_thread = threading.Thread(target=run_assistant)
    assistant_thread.start()
    return jsonify({'response': 'Voice Assistant started'})


if __name__ == "__main__":
    app.run(debug=True)