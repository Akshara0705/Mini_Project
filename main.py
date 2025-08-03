import speech_recognition as sr
import os
import pyttsx3
import ctypes
import requests
import webbrowser
import datetime
from dateutil import parser
import openai
import cohere
from config import apikey
import threading
from fuzzywuzzy import process
import Levenshtein
import subprocess
import winreg
import ctypes
import pathlib
import win32com.client
import psutil
import nltk
from nltk.corpus import wordnet
import re
from word2number import w2n
import pyautogui
from pywinauto import Application, Desktop
import pygetwindow as gw
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import glob


chatStr = ""
# Download WordNet data (only needed once)
#nltk.download('wordnet')
driver = None
#interrupt_flag = threading.Event()
def reset_chat():
    global chatStr
    chatStr = ""
    say("Chat history has been reset.")

def say(text):
    global engine
    print(text)
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.say({text})
    engine.runAndWait()


def takeCommand():
    global chatStr
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            #say("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            say(f"User said: {query}")
            chatStr += f"You: {query}\n"
            return query
        except Exception as e:
            return takeCommand()


'''def listen_for_interrupt():
    global interrupt_flag
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)  # Adjust for ambient noise
        say("Listening for interrupt command...")
        while not interrupt_flag.is_set():
            try:
                audio = recognizer.listen(source, timeout=3, phrase_time_limit=5)  # Adjusted timeout and phrase_time_limit
                query = recognizer.recognize_google(audio, language="en-in").lower()
                say(f"Recognized command: {query}")
                if "stop" in query:
                    interrupt_flag.set()  # Set the interrupt flag
                    engine.stop()  # Stop the TTS engine
                    say("Audio interrupted.")
                    break
            except sr.UnknownValueError:
                say("Could not understand the command.")
            except sr.RequestError as e:
                say(f"Could not request results; {e}")
                break
            except Exception as e:
                say(f"An unexpected error occurred: {e}")'''

'''def play_and_listen(text):
    global engine, interrupt_flag

    interrupt_flag.clear()  # Clear the interrupt flag

    # Start the say function in a separate thread
    tts_thread = threading.Thread(target=say, args=(text,))
    tts_thread.start()

    # Listen for interrupt command
    listen_thread = threading.Thread(target=listen_for_interrupt)
    listen_thread.start()

    # Wait for the TTS thread to finish
    tts_thread.join()
    # Set the interrupt flag in case TTS finished before listening
    interrupt_flag.set()
    # Wait for the listening thread to finish
    listen_thread.join()'''

'''def takeCommand():
    global chatStr
    try:
        # Taking input from the user through text
        query = input("Please enter your command: ")
        say(f"User said: {query}")
        chatStr += f"You: {query}\n"
        return query
    except Exception as e:
        return "Some Error Occurred. Sorry"'''




def chat(query):
    global chatStr
    co = cohere.Client(apikey)

    chatStr += f"User: {query} \n Assistant: "

    try:
        response = co.generate(
            model='command-xlarge-nightly',
            prompt=chatStr,
            temperature=0.7,
            max_tokens=256
        )


        response_text = response.generations[0].text.strip()


        print("Assistant: ")
        formatted_response = f"{response_text}"
        say(response_text)

        chatStr += f"{response_text}\n"

        return response_text
    except Exception as e:
        error_message = f"An error occurred: {e}"
        say(error_message)
        return error_message


def format_website_url(website_name):
    """
    Format the website URL by attempting to add 'https://' and 'http://'
    to the given website name. Assumes '.com' if no domain is provided.

    Args:
    - website_name (str): The name of the website (e.g., 'openai' or 'perplexity.ai').

    Returns:
    - str: The formatted URL if the website is reachable, otherwise an empty string.
    """
    possible_schemes = ['https://', 'http://']

    # Check if website_name includes a domain extension
    if '.' not in website_name:
        website_name += '.com'

    for scheme in possible_schemes:
        url = f"{scheme}{website_name}"
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                return url
        except requests.exceptions.RequestException:
            continue

    # If none of the combinations worked, return an empty string
    return ""

def get_drives():
    # This function returns a list of available drives (Windows specific)
    drives = []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        if bitmask & 1:
            drives.append(letter + ':\\')
        bitmask >>= 1
    return drives

def search_folders(folder_name):
    normalized_folder_name = normalize_string(folder_name)
    drives = get_drives()
    matching_folders = []

    for drive in drives:
        for root, dirs, _ in os.walk(drive):
            for d in dirs:
                if normalized_folder_name in normalize_string(d):
                    matching_folders.append(os.path.join(root, d))

    return matching_folders

def search_file(filename, extensions=None):
    file_paths = []
    drives = [f"{d}:\\" for d in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if os.path.exists(f"{d}:\\")]
    filename_pattern = re.compile(re.escape(filename), re.IGNORECASE)

    for drive in drives:
        for root, dirs, files in os.walk(drive):
            for name in files:
                if re.search(filename_pattern, name):
                    if extensions:
                        if any(name.lower().endswith(ext) for ext in extensions):
                            file_paths.append(os.path.join(root, name))
                    else:
                        file_paths.append(os.path.join(root, name))
    return file_paths

def normalize_string(s):
    return re.sub(r'[^A-Za-z0-9]', '', s).lower()

def get_cohere_response(api_key, query):
    co = cohere.Client(api_key)

    try:
        response = co.generate(
            model='command-r-plus',
            prompt=query,
            max_tokens=50
        )

        return response.generations[0].text.strip()
    except cohere.CohereError as e:
        return f"An error occurred: {e}"

def close_website_tab(tab_title):
    say("Attempting to close website tab...")
    windows = gw.getWindowsWithTitle(" - Google Chrome")
    for window in windows:
        say(f"Window title: {window.title}")
        if tab_title.lower() in window.title.lower():
            say(f"Found window with title: {window.title}")
            window.activate()
            pyautogui.hotkey('ctrl', 'w')
            say(f"Closed tab with title: {tab_title}")
            return True
    say(f"Could not find the website: {tab_title}")
    return False


def close_application(query):
    """Function to close an application by name"""
    # Extract the website name from the query
    match = re.search(r'close\s+(.*?)\s+(website|app)', query.lower())
    if match:
        app_name = match.group(1).strip().lower()
    else:
        say("Could not extract the website or app name from the query.")
        return
    found = False
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if app_name.lower() in proc.info['name'].lower():
                os.system(f'taskkill /F /PID {proc.info["pid"]}')
                say(f"Closed {proc.info['name']} with PID {proc.info['pid']}")
                found = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    # Additional check for UWP apps
    if not found:
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                cmdline = ' '.join(proc.cmdline()).lower()
                if app_name.lower() in cmdline:
                    os.system(f'taskkill /F /PID {proc.info["pid"]}')
                    say(f"Closed UWP app {proc.info['name']} with PID {proc.info['pid']}")
                    found = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

    if not found:
        say(f"Could not find the application: {app_name}")
    return found


def is_agreement_word(word):
    synonyms = set()
    for syn in wordnet.synsets(word):
        for lemma in syn.lemmas():
            synonyms.add(lemma.name().lower())
    agreement_words = {"yes", "yeah", "yup", "go", "sure", "okay", "ok", "alright", "affirmative", "indeed"}
    return not synonyms.isdisjoint(agreement_words)

def get_program_files_paths():
    """Retrieve paths from Windows Registry for common program files locations."""
    program_files_paths = []

    try:
        # Open registry key for 64-bit programs
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths")
        index = 0
        while True:
            sub_key_name = winreg.EnumKey(key, index)
            sub_key = winreg.OpenKey(key, sub_key_name)
            path, _ = winreg.QueryValueEx(sub_key, "")
            program_files_paths.append(path)
            index += 1
    except FileNotFoundError:
        pass  # Handle if the registry key doesn't exist or is inaccessible
    except WindowsError:
        pass  # Handle any other Windows-specific errors

    return program_files_paths

def close_website(query):
    # Extract the website name from the query
    match = re.search(r'close\s+(.*?)\s+(website|app)', query.lower())
    if match:
        target_name = match.group(1).strip().lower()
    else:
        say("Could not extract the website or app name from the query.")
        return

    found = False

    # Get all open windows
    windows = gw.getAllWindows()

    for window in windows:
        # say window titles for debugging
        print(window.title.lower())

        # Check if the target name is in the window title
        if target_name in window.title.lower():
            # Activate the window
            window.activate()

            # Wait for a moment
            time.sleep(0.5)

            # Send the hotkey to close the tab or window
            pyautogui.hotkey('ctrl', 'w')

            say(f"Closed {target_name} website.")
            found = True
            break

    if not found:
        say(f"Could not find the website or app: {target_name}")

def open_folder_location(folder_path):
    if os.path.exists(folder_path):
        # This command is for Windows, adapt as needed for other OS (e.g., `open` for macOS, `xdg-open` for Linux)
        subprocess.Popen(f'explorer /select,"{folder_path}"')
        say(f"Opened the folder location: {folder_path}")
    else:
        say(f"Folder '{folder_path}' does not exist.")

def open_path_location_of_folder(query):
    folder_name = query.replace('open path location of folder', '').strip()
    if folder_name:
        say(f"Searching for folder: {folder_name}")
        folder_paths = search_folders(folder_name)
        if not folder_paths:
            say(f"No folders found with the name {folder_name}")
        elif len(folder_paths) == 1:
            open_folder_location(folder_paths[0])
        else:
            say(f"Multiple folders found with the name {folder_name}. Please select one.")
            for i, path in enumerate(folder_paths):
                print(f"{i + 1}: {path}")
            try:
                say("Please type the number of the folder whose location you want to open.")
                choice = int(input("Enter the number of the folder whose location you want to open: "))
                if 1 <= choice <= len(folder_paths):
                    open_folder_location(folder_paths[choice - 1])
                else:
                    say("Invalid choice.")
            except Exception as e:
                say(e)
                say("Sorry, I couldn't understand your choice.")

def find_uwp_app(app_name):
    """Find UWP app using PowerShell."""
    script = f"""
    Get-StartApps | Where-Object {{ $_.Name -like "*{app_name}*" }} | Select-Object -ExpandProperty AppID
    """
    process = subprocess.Popen(
        ["powershell", "-Command", script],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()
    if stdout:
        app_id = stdout.decode().strip()
        if app_id:
            return app_id
    return None

def find_uwp_app(app_name):
    """Find UWP app using PowerShell."""
    script = f"""
    Get-StartApps | Where-Object {{ $_.Name -like "*{app_name}*" }} | Select-Object -ExpandProperty AppID
    """
    process = subprocess.Popen(
        ["powershell", "-Command", script],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()
    if stdout:
        app_id = stdout.decode().strip()
        if app_id:
            return app_id
    return None

def find_app(app_name):
    """Find the full path of an installed application by name."""
    # Check for UWP app
    app_user_model_id = find_uwp_app(app_name)
    if app_user_model_id:
        return f"explorer.exe shell:AppsFolder\\{app_user_model_id}"

    # Common directories to search for executables
    paths = [
        "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs",
        os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs"),
        os.path.expanduser("~\\Desktop"),  # Add desktop shortcuts
        os.path.expanduser("~\\AppData\\Local\\Microsoft\\WindowsApps"),  # Add Windows Apps
        "C:\\Program Files",
        "C:\\Program Files (x86)",
    ]

    # Search for the application executable in the directories
    for path in paths:
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.lower().startswith(app_name.lower()) and file.endswith('.exe'):
                    return os.path.join(root, file)
    return None



'''def open_website(url):
    global driver
    say("Opening website...")
    chrome_options = Options()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get(url)
    say("Website opened.")
    return driver'''

def open_app(query):
    match = re.search(r'open\s+(.*?)\s+app', query.lower())
    if match:
        app_name = match.group(1).strip()
    else:
        say("Could not extract app name.")
    app_path = find_app(app_name)
    if app_path:
        say(f"Opening {app_name} ...")
        subprocess.Popen(app_path, shell=True)
    else:
        say(f"Sorry, I couldn't find the application you mentioned.")
        say(f"Search for website instead?")
        print('Listening...')
        querry = takeCommand()
        if is_agreement_word(querry):
            #url=format_website_url(app_name)
            #open_website(url)
            open_website(query)
        else:
            say(f"okay")
    return

def open_file(file_path):
    try:
        say(f"Opening file: {file_path}")
        os.startfile(file_path)
        say(f"Opening file: {file_path}")
    except Exception as e:
        say(e)
        say("Sorry, I couldn't open the file.")

def open_file_location(file_path):
    folder_path = os.path.dirname(file_path)
    try:
        say(f"Opening folder: {folder_path}")
        os.startfile(folder_path)
        say(f"Opening folder: {folder_path}")

        app = Application().connect(path="explorer.exe", timeout=10)
        app.top_window().set_focus()
        app.top_window().maximize()
    except Exception as e:
        say(e)
        say("Sorry, I couldn't open the folder.")

def open_document_location(document_name):
    say(f"Searching for document: {document_name}")
    document_paths = search_documents(document_name)
    if not document_paths:
        say(f"No documents found with the name {document_name}")
    elif len(document_paths) == 1:
        open_file_location(document_paths[0])
    else:
        say(f"Multiple documents found with the name {document_name}. Please select one.")
        for i, path in enumerate(document_paths):
            print(f"{i + 1}: {path}")
        try:
            say("Please type the number of the document whose folder you want to open.")
            choice = int(input("Enter the number of the document whose folder you want to open: "))
            if 1 <= choice <= len(document_paths):
                open_file_location(document_paths[choice - 1])
            else:
                say("Invalid choice.")
        except Exception as e:
            say(e)
            say("Sorry, I couldn't understand your choice.")

def open_folder(folder_name):
    matching_folders = search_folders(folder_name)

    if not matching_folders:
        say(f"No folders found matching '{folder_name}'")
        return

    if len(matching_folders) == 1:
        folder_to_open = matching_folders[0]
    else:
        say(f"Multiple folders found matching '{folder_name}':")
        for idx, folder in enumerate(matching_folders, start=1):
            say(f"{idx}. {folder}")

        say("Please say the number of the folder you want to open.")
        choice = takeCommand()
        try:
            choice_num = w2n.word_to_num(choice)
        except ValueError:
            choice_num = -1
        if 1 <= choice_num <= len(matching_folders):
            folder_to_open = matching_folders[choice_num - 1]
        else:
            say("Invalid choice. Operation cancelled.")
            return

    # Open the folder
    if os.path.exists(folder_to_open):
        # This command is for Windows, adapt as needed for other OS (e.g., `open` for macOS, `xdg-open` for Linux)
        subprocess.Popen(f'explorer "{folder_to_open}"')

        # Give some time for the folder to open
        time.sleep(2)

        # Extract folder name from path to get the window title
        folder_title = os.path.basename(folder_to_open)

        # Get the window and bring it to the front
        windows = gw.getWindowsWithTitle(folder_title)
        if windows:
            windows[0].activate()
            windows[0].maximize()
            say(f"Opened and focused folder: {folder_title}")
        else:
            say(f"Folder '{folder_title}' opened but could not bring to front.")
    else:
        say(f"Folder '{folder_to_open}' does not exist.")

def list_open_windows():
    windows = gw.getAllTitles()
    if not windows:
        say("No open windows found.")
        return
    say("Open windows:")
    for window in windows:
        if window:
            print(window)
    say("Here are the open windows.")

def switch_to_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)
        if not window:
            say(f"No window found with the title: {window_title}")
            return False
        window[0].activate()
        say(f"Switched to window: {window_title}")
        return True
    except Exception as e:
        say(e)
        say("Sorry, I couldn't switch to the window.")
        return False

def minimize_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)
        if not window:
            say(f"No window found with the title: {window_title}")
            return False
        window[0].minimize()
        say(f"Minimized window: {window_title}")
        return True
    except Exception as e:
        say(e)
        say("Sorry, I couldn't minimize the window.")
        return False

def maximize_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)
        if not window:
            say(f"No window found with the title: {window_title}")
            return False
        window[0].maximize()
        say(f"Maximized window: {window_title}")
        return True
    except Exception as e:
        say(e)
        say("Sorry, I couldn't maximize the window.")
        return False

def close_window(window_title):
    try:
        window = gw.getWindowsWithTitle(window_title)
        if not window:
            say(f"No window found with the title: {window_title}")
            return False
        window[0].close()
        say(f"Closed window: {window_title}")
        return True
    except Exception as e:
        print(e)
        say("Sorry, I couldn't close the window.")
        return False

def open_website(search_url):
    """Open an application or website based on the query."""
    #match = re.search(r'open\s+(.*?)\s+(website|app)', query.lower())
    #if match:
        #website_name = match.group(1).strip()
    #else:
        #say("Could not extract website name from the query.")
        #return
    #app_path = find_app(app_name)
    #if app_path:
        #say(f"Opening {app_name} ...")
        #subprocess.Popen(app_path, shell=True)
    #else:
        #say(f"Sorry, I couldn't find the application you mentioned. Let's search for website")
    #search_url = f"https://{website_name}.com"
    say(f"Opening...")
    webbrowser.open(search_url)
    return

def open_website_in_new_window(website):
    say(f"Opening {website} website...")
    try:
        if os.name == 'nt':
            subprocess.Popen(['start', 'chrome', '--new-window', website], shell=True)
    except Exception as e:
        say(f"Failed to open {website} in a new window. Error: {e}")
        webbrowser.open_new_tab(website)

def close_website(query):
    # Extract the website name from the query
    match = re.search(r'close\s+(.*?)\s+website', query.lower())
    if match:
        website_name = match.group(1).strip()
    else:
        say("Could not extract website name from the query.")
        return

    # List of common web browser executable names
    browsers = ["chrome.exe", "firefox.exe", "msedge.exe", "opera.exe"]

    found = False

    # Iterate over all running processes
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            # Check if the process is a browser
            if proc.info['name'] in browsers:
                # Check if the browser's command line contains the website name
                if any(website_name in arg.lower() for arg in proc.info['cmdline']):
                    proc.terminate()  # Terminate the process
                    found = True
                    say(f"Closed {website_name} website.")
                    break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    if not found:
        say(f"Could not find the website: {website_name}")

def ai(query, api_key):
    response = get_cohere_response(api_key, query)


    os.makedirs("Openai", exist_ok=True)


    filename = f"Openai/{''.join(query.split('intelligence')[1:]).strip()}.txt"
    with open(filename, "w") as f:
        f.write(response)

   # say(response)


if __name__ == '__main__':
    say('Welcome')
    while True:
        print("Listening...")
        query = takeCommand()
        #say(query)
        # TODO: add more sites

        #if "open music" in query:
            #musicPath = "C:\\Users\Sagarika\Downloads\\flow-211881.mp3"
            #os.system(f'start "" "{musicPath}"')

        if "the time" in query.lower():
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Time is {hour} hours {min} minutes")

        elif "Using artificial intelligence".lower() in query.lower():
            """ai_thread = threading.Thread(target=ai, args=(query, apikey))
            ai_thread.start()"""
            ai(query,apikey)

            # using artificial intelligence write an essay on fruits

        elif 'open' in query and 'in new window' in query and 'website' in query:
            '''match = re.search(r'open\s+(.*?)\s+website', query.lower())
            if match:
                app_name = match.group(1).strip()
            else:
                say("Could not extract app name.")
            url = format_website_url(app_name)
            #open_website(url)
            open_website_in_new_window(url)'''
            website = query.replace('open', '').replace('website', '').replace('in new window', '').strip()
            url = format_website_url(website)
            open_website_in_new_window(url)


        elif 'open' in query and 'in new window' in query and '.' in query:
            '''match = re.search(r'open\s+(.*?)\s+in new window', query.lower())
            if match:
                app_name = match.group(1).strip()
            else:
                say("Could not extract app name.")
            url = format_website_url(app_name)
            # open_website(url)
            open_website_in_new_window(url)'''
            website = query.replace('open', '').replace('in new window', '').strip()
            url = format_website_url(website)
            open_website_in_new_window(url)

        elif "open" in query.lower() and "website" in query.lower():
            match = re.search(r'open\s+(.*?)\s+website', query.lower())
            if match:
                app_name = match.group(1).strip()
            else:
                say("Could not extract app name.")
            url=format_website_url(app_name)
            #open_website(url)
            #open_website(query)

            open_website_in_new_window(url)


        elif "open" in query.lower() and "." in query.lower():
            app_name = query.lower().replace("open ", "").strip()
            url = format_website_url(app_name)
            #open_website(url)
            #open_website(query)
            open_website_in_new_window(url)

        elif "open" in query.lower() and "app" in query.lower():
            open_app(query)
        elif "close" in query.lower() and "app" in query.lower():
            close_application(query)

        elif 'open file' in query.lower():
            filename = query.lower().replace('open file', '').strip()
            extensions = None
            if 'image' in query.lower():
                extensions = ['.png', '.jpg', '.jpeg']
                filename = filename.lower().replace(' image', '').strip()
            elif 'photo' in query.lower():
                extensions = ['.png', '.jpg', '.jpeg']
                filename = filename.lower().replace(' photo', '').strip()
            elif 'pdf' in query.lower():
                extensions = ['.pdf']
                filename = filename.lower().replace(' pdf', '').strip()
            elif 'text' in query.lower():
                extensions = ['.txt']
                filename = filename.lower().replace(' txt', '').strip()
            elif 'ppt' in query.lower():
                extensions = ['.ppt', '.pptx']
                filename = filename.lower().replace(' ppt', '').strip()
            elif 'mp3' in query.lower():
                extensions = ['.mp3']
                filename = filename.lower().replace(' mp3', '').strip()
            elif 'mp4' in query.lower():
                extensions = ['.mp4']
                filename = filename.lower().replace(' mp4', '').strip()

            if filename:
                say(f"Searching for file: {filename}")
                file_paths = search_file(filename, extensions)
                if not file_paths:
                    say(f"No files found with the name {filename}")
                elif len(file_paths) == 1:
                    open_file(file_paths[0])
                else:
                    say(f"Multiple files found with the name {filename}. Please select one.")
                    for i, path in enumerate(file_paths):
                        print(f"{i + 1}: {path}")
                    try:
                        say("Please say the number of the file you want to open.")
                        choice = takeCommand()
                        try:
                            choice_num = w2n.word_to_num(choice)
                        except ValueError:
                            choice_num = -1
                        if 1 <= choice_num <= len(file_paths):
                            open_file(file_paths[choice_num - 1])
                        else:
                            print("Invalid choice")
                    except Exception as e:
                        say(e)
                        say("Sorry, I couldn't understand your choice.")

        elif 'open path location of file' in query.lower():
            filename = query.replace('open path location of file', '').strip()
            if filename:
                say(f"Searching for file: {filename}")
                file_paths = search_file(filename)
                if not file_paths:
                    say(f"No files found with the name {filename}")
                elif len(file_paths) == 1:
                    open_file_location(file_paths[0])
                else:
                    say(f"Multiple files found with the name {filename}. Please select one.")
                    for i, path in enumerate(file_paths):
                        print(f"{i + 1}: {path}")
                    try:
                        say("Please type the number of the file whose folder you want to open.")
                        choice = int(input("Enter the number of the file whose folder you want to open: "))
                        if 1 <= choice <= len(file_paths):
                            open_file_location(file_paths[choice - 1])
                        else:
                            say("Invalid choice.")
                    except Exception as e:
                        say(e)
                        say("Sorry, I couldn't understand your choice.")

        if 'open path location of folder' in query:
            open_path_location_of_folder(query)


        elif 'open folder' in query.lower():
            folder_name = query.replace('open folder', '').strip()
            open_folder(folder_name)

        elif 'switch to tab' in query.lower():
            window_title = query.replace('switch to tab', '').strip()
            if window_title:
                result = switch_to_window(window_title)
                if result:
                    say("Switched to the window successfully.")
                else:
                    say("Failed to switch to the window.")

        elif 'minimise tab' in query.lower() or 'minimise window' in query.lower():
            window_title = query.replace('minimise tab', '').replace('minimise window', '').strip()
            if window_title:
                result = minimize_window(window_title)
                if result:
                    say("Minimized the window successfully.")
                else:
                    say("Failed to minimize the window.")

        elif 'maximize tab' in query.lower() or 'maximize window' in query.lower():
            window_title = query.replace('maximize tab', '').replace('maximize window', '').strip()
            if window_title:
                result = maximize_window(window_title)
                if result:
                    say("Maximized the window successfully.")
                else:
                    say("Failed to maximize the window.")

        elif 'close tab' in query.lower() or 'close window' in query.lower():
            window_title = query.lower().replace('close tab', '').replace('close window', '').strip()
            print(window_title)
            if window_title:
                result = close_window(window_title)
                if result:
                    say("Closed the window successfully.")
                else:
                    say("Failed to close the window.")

        elif 'show all tabs' in query.lower():
            list_open_windows()

            



        elif "terminate".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            reset_chat()

        else:
            print("Chatting...")
            #chat(query)
