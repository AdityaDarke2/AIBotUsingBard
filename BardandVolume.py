from bardapi import Bard
import os
import sys
import speech_recognition as sr
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import webbrowser


os.environ["_BARD_API_KEY"] = "agjgoTmqWBIHDyIawZezroTinvdpY6qHhuKytIh_9Ysp1qTWqqFghPOZwG9MOkllwQu52A."

recognizer = sr.Recognizer()

chrome_file_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

def blockPrint():
    sys.stdout = open(os.devnull, 'w')

def enablePrint():
    sys.stdout = sys.__stdout__

def listen_and_transcribe():

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)  
        print("Listening...")

        try:
            audio = recognizer.listen(source)
            text = recognizer.recognize_google(audio)
            
            return text.lower()  
        
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand what you said.")
            return 
        
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return None        

def change_volume_windows(volume_percentage):
    
    volume_percentage = max(0, min(100, volume_percentage))

    devices = AudioUtilities.GetSpeakers()

    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    scalar_volume = volume_percentage / 100.0

    volume.SetMasterVolumeLevelScalar(scalar_volume, None)

def check_shutdown_command(spoken_text):

    if spoken_text == "shut down":
        return True
    

    elif spoken_text == "shutdown":
        return True
    
    
    return False

def check_wake_word(spoken_text):

    if "Iris" in spoken_text :
        return True
    

    elif "iris" in spoken_text:
        return True
    

    return False

def get_volume_percentage_from_voice(command):


   try:
        if command:
            words = command.split()
            for i, word in enumerate(words):
                if word == 'to' and i < len(words) - 1:
                    try:
                        volume_percentage = int(words[i + 1])
                        return volume_percentage
                    except ValueError:
                        pass
        return None


   except sr.UnknownValueError:
            print("Could not understand audio")
            return None
   

   except sr.RequestError as e:
        print("Could not request results; {0}".format(e))
        return None

def common_questions():
        

        if 'open youtube' in spoken_text:
             print('There you Go!')
             webbrowser.get(chrome_file_path).open("youtube.com")
        

        elif 'open google' in spoken_text:
             print('There you Go!')
             webbrowser.get(chrome_file_path).open("google.com")


        elif 'weather' in spoken_text:
            query = spoken_text.replace('iris', ' ')
            print("Iris' Response: " + Bard().get_answer(str(query))['content'])
 

        elif 'name' in spoken_text:
            print('My name is Iris. I was developed by a group of students studying at XXXXX Engineering College, XXXXX')


        elif 'developed you' in spoken_text:
            print('I was developed by a group of students studying at XXXXX Engineering College, XXXXX')


        elif 'about yourself' in spoken_text:
            print('My Name is Iris. I was developed by a group of students studying at XXXXX Engineering College, XXXXX')

while True:
    spoken_text = listen_and_transcribe()
    volume_percentage = get_volume_percentage_from_voice(spoken_text)

    if spoken_text:
        print(f"You said: {spoken_text}")

        common_questions()

        if check_shutdown_command(spoken_text):

            print("I am on Sleep mode but you can wake me up by saying Iris")
            blockPrint()

        elif check_wake_word(spoken_text):
            enablePrint()
            query = spoken_text.replace('iris', ' ')
            print("Iris' Response: " + Bard().get_answer(str(query))['content'])

        elif volume_percentage is not None:
                change_volume_windows(volume_percentage)
                print(f"Volume changed to {volume_percentage}%")

        

