import speech_recognition as sr
import webbrowser
import time
import AppOpener
from gtts import gTTS
from pygame import mixer
import ctypes
import pyautogui
import winshell
import threading
from pynput.keyboard import Listener
import win32api
from win32con import VK_MEDIA_PLAY_PAUSE, KEYEVENTF_EXTENDEDKEY
import pyvolume
import json

def recognize_speech_from_mic(recognizer:sr.Recognizer, microphone:sr.Microphone, timeout:int=None):

    with microphone as source:
        try:

            audio = recognizer.listen(source, phrase_time_limit=timeout)
            response = {"success": True, "error": None, "transcription": None}
            response["transcription"] = recognizer.recognize_google(audio)

        except sr.RequestError:

            response["success"] = False
            response["error"] = "API unavailable"

        except sr.UnknownValueError:

            response["error"] = "Unable to recognize speech"

        except:

            response["transcription"] = ""

    print(response["transcription"])
    return response

class Voice:

    def __init__(self, response, recognizer, microphone):

        self.response = response
        self.transcription:str = None
        self.reminder_thread = threading.Thread()
        self.timer_thread = threading.Thread()
        self.recognizer = recognizer
        self.microphone = microphone
        self.playing = False
        self.welcome = False
        self.checked_reminders = False
        self.checked_timers = False
        mixer.init()

    def welcome_func(self):

        if self.checked_reminders and self.checked_timers:
            tts = gTTS(f"Click the right control key to start!", tld="us")
            tts.save("./save.mp3")
            mixer.music.load('./save.mp3')
            mixer.music.play()

    def sleep(self):

        tts = gTTS(f'locking the computer.', tld="us")
        tts.save('sleep.mp3')
        mixer.music.load("./sleep.mp3")
        mixer.music.play()
        ctypes.windll.user32.LockWorkStation()

    def search_for(self, search):

        suffix_list = (".com", ".to", ".org", ".in", ".io", "us")

        if search.endswith(suffix_list):
            url = f"https://{search}"
        else:
            search = search.replace(" ", "+")
            url = f"https://www.google.com/search?-b-d&q={search}"

        webbrowser.open_new_tab(url)
        tts = gTTS(f'Searching for {search}', tld="us")
        tts.save('search.mp3')
        mixer.music.load("./search.mp3")
        mixer.music.play()
    
    def open(self, prompt):

        if prompt.lower() == "anime":
            webbrowser.open_new_tab("https://hianime.to/user/continue-watching")
            tts = gTTS(f'opening hianime.to', tld="us")
            tts.save('hianime.mp3')
            mixer.music.load("./hianime.mp3")
            return mixer.music.play()

        try:
            AppOpener.open(prompt, match_closest=True)
            tts = gTTS(f'Opening {prompt}', tld="us")
            tts.save('open.mp3')
            mixer.music.load("./open.mp3")
            mixer.music.play()
        except:
            print(f"Couldn't find {prompt}")

    def type(self, prompt):

        pyautogui.typewrite(prompt)

    def clear_bin(self):

        winshell.recycle_bin().empty(confirm=False)
        tts = gTTS(f'Emptying the recycle bin', tld="us")
        tts.save('empty.mp3')
        mixer.music.load("./empty.mp3")
        mixer.music.play()
    
    def close(self, prompt):

        try:
            AppOpener.close(prompt, match_closest=True)
            tts = gTTS(f'Closing {prompt}', tld="us")
            tts.save('close.mp3')
            mixer.music.load("./close.mp3")
            mixer.music.play()
        except:
            print(f"Couldn't find {prompt}")

    def play(self):

        if self.playing == False:
            win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, 0)
            self.playing = True
            
    
    def pause(self):

        if self.playing == True:
            win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, 0)
            self.playing = False

    def get_new_response(self, timeout):

        self.response = recognize_speech_from_mic(self.recognizer, self.microphone, timeout)
        self.transcription = self.response["transcription"]

    def volume_control(self):

        single_digits = {'zero': 0,
                        'one': 1,
                        'two': 2,
                        'three': 3,
                        'four': 4,
                        'five': 5,
                        'six': 6,
                        'seven': 7,
                        'eight': 8,
                        'nine': 9
                    }
        
        idk_list = self.transcription.lower().split(" ")

        last_num = None

        try:
            last_num = int(idk_list[len(idk_list)-1])
        except:
            for word in idk_list:
                try:
                    last_num = int(word)
                    break
                except:
                    pass
            try:
                for i in single_digits.keys():

                    if i in idk_list:
                        print(1)
                        last_num = single_digits[i]

            except Exception as e:
                print(e)
                pass

        if last_num == None:
            return "Couldn't get the volume."

        pyvolume.custom(last_num)

    def time_to_seconds(self, time_str):

        units_in_seconds = {
            'hour': 3600,
            'hours': 3600,
            'minute': 60,
            'minutes': 60,
            'second': 1,
            'seconds': 1,
}
        
        parts = time_str.split()

        if len(parts) != 2:
            raise ValueError("Invalid input format. Please provide a string in the format 'number unit'.")
        
        number = float(parts[0])
        unit = parts[1].lower()
        
        if unit in units_in_seconds:
            return number * units_in_seconds[unit]
        else:
            raise ValueError(f"Invalid time unit: {unit}")

    def save_remind(self, sections):

        about = sections[1]
        string_time = sections[2]

        try:
            in_sec = self.time_to_seconds(string_time)
        except Exception as e:
            tts = gTTS(f'Could you please say the time in a proper format.', tld="us")
            tts.save('imnotaaiyouidioto.mp3')
            mixer.music.load("./imnotaaiyouidioto.mp3")
            mixer.music.play()
        
        remind_in = time.time() + in_sec

        with open("./data.json", "r+") as e:

            data = json.load(e)
            data["reminders"].append([sections[0], about, remind_in])
            e.seek(0)
            json.dump(data, e)
            e.truncate()

            if sections[0] == "remind me about ":
                tts = gTTS(f'Alright! I will remind you about {about} in {string_time}', tld="us")
            else:
                tts = gTTS(f'Alright! I will remind you to {about} in {string_time}', tld="us")

            tts.save('imnotaaiyouidioto.mp3')
            mixer.music.load("./imnotaaiyouidioto.mp3")
            mixer.music.play()

    def save_timer(self, string_time):

        with open("./data.json", "r+") as e:

            data = json.load(e)
            data["timers"].append(time.time()+self.time_to_seconds(string_time))
            e.seek(0)
            json.dump(data, e)
            e.truncate()

        tts = gTTS(f'Alright! timer for {string_time} has been set.', tld="us")
        tts.save('timerimnotai.mp3')
        mixer.music.load("./timerimnotai.mp3")
        mixer.music.play()

    def remove_reminder(self, all:bool=False):

        with open("./data.json", "r+") as e:

            data = json.load(e)

            if all == True:
                tts = gTTS(f'Removed all reminders!', tld="us")
                data["reminders"] = []
            else:
                tts = gTTS(f"Removed the last reminder" , tld = "us")
                data["reminders"].pop()

            tts.save('reminders_removed.mp3')
            mixer.music.load("./reminders_removed.mp3")
            e.seek(0)
            json.dump(data, e)
            e.truncate()
            mixer.music.play()

    def remove_timer(self, all:bool=False):

        with open("./data.json", "r+") as e:

            data = json.load(e)
            if all == True:  
                tts = gTTS(f'Removed all timers!', tld="us")
                data["timers"] = []
            else:
                tts = gTTS(f"Removed the last timer", tld = "us")
                data["timers"].pop()

            tts.save('reminders_removed.mp3')
            mixer.music.load("./reminders_removed.mp3")
            e.seek(0)
            json.dump(data, e)
            e.truncate()
            mixer.music.play()

    def check(self):

        mixer.music.load("./BEEP MF.mp3")
        (threading.Thread(target=mixer.music.play, daemon=True)).start()

        self.get_new_response(timeout=5)

        if self.response["error"]:
            print(self.response["error"])

        if "go to sleep" in self.transcription.lower() or "lock the pc" in self.transcription.lower():
            self.sleep()

        elif self.transcription.lower().startswith("search for"):
            self.search_for(self.transcription.lower().removeprefix("search for "))

        elif self.transcription.lower().startswith("open"):
            self.open(self.transcription.lower().removeprefix("open "))

        elif self.transcription.lower().startswith("type"):
            self.type(self.transcription.removeprefix("type "))

        elif "clear recycle bin" in self.transcription or "empty recycle bin" in self.transcription:
            self.clear_bin()

        elif self.transcription.lower().startswith("close"):
            self.close(self.transcription.lower().removeprefix("close "))

        elif "pause" in self.transcription.lower():
            self.pause()

        elif "play" in self.transcription.lower() or "resume" in self.transcription.lower():
            self.play()

        elif "volume" in self.transcription.lower():
            self.volume_control()

        elif self.transcription.lower().startswith("remind me about"):
            sections = self.transcription.lower().split("remind me about ")
            end_part = sections[1].split(" ")

            time = f"{end_part[-2]} {end_part[-1]}"
            sections[1] = sections[1].removesuffix(f" {end_part[-3]} {time}")
            sections[0] = "about"

            sections.append(time)
            self.save_remind(sections)

        elif self.transcription.lower().startswith("remind me to"):
            sections = self.transcription.lower().split("remind me to ")
            end_part = sections[1].split(" ")

            time = f"{end_part[-2]} {end_part[-1]}"
            sections[1] = sections[1].removesuffix(f" {end_part[-3]} {time}")
            sections[0] = "to"

            sections.append(time)
            self.save_remind(sections)

        elif self.transcription.lower().startswith("set a timer of"):
            time = self.transcription.lower().removeprefix("set a timer of ")

            self.save_timer(time)

        elif self.transcription.lower().startswith("remove the last"):
            to_remove = self.transcription.lower().removeprefix("remove the last ")
            if to_remove == "reminder":
                print(1)
                self.remove_reminder()
            if to_remove == "timer":
                self.remove_timer()

        elif self.transcription.lower().startswith("remove all "):
            to_remove = self.transcription.lower().removeprefix("remove all ")
            if to_remove == "reminders":
                self.remove_reminder(all=True)
            if to_remove == "timers":
                self.remove_timer(all=True)

    def run(self):
        if not self.welcome:
            self.welcome_func()
            self.welcome = True
            print("loaded")

        def keypress(key):
            try:
                key = key.name
                if key == "ctrl_r":
                    self.check()
            except:
                pass

        with Listener(on_press=keypress) as listener:
            listener.join()
    
    def reminder_loop(self):
        mixer.init()
        while True:
            with open("./data.json", "r+") as e:
                data = json.load(e)

                reminders = data["reminders"]

                i = 0
                for reminder in reminders:
                    if float(reminder[2]) <= time.time():
                        tts = gTTS(f'Reminding you {reminder[0]} {reminder[1]}', tld="us")
                        tts.save('reminder.mp3')
                        mixer.music.load("./reminder.mp3")
                        mixer.music.play()
                        del reminders[i]
                    i += 1

                data["reminders"] = reminders
                e.seek(0)
                json.dump(data, e)
                e.truncate()
                self.checked_reminders = True
        
            time.sleep(5)

    def timer_loop(self):
        mixer.init()
        while True:
            with open("./data.json", "r+") as e:
                data = json.load(e)
                timers = data["timers"]
                i = 0
                for timer in timers:
                    if float(timer) <= time.time():
                        mixer.music.load("./BEEP BEPP.mp3")
                        mixer.music.play()
                        del timers[i]
                    i += 1
                
                data["timers"] = timers
                e.seek(0)
                json.dump(data, e)
                e.truncate()
                self.checked_timers = True

            time.sleep(5)

    def main(self):
        mixer.init()

        self.reminder_thread = threading.Thread(target=self.reminder_loop, daemon=True)
        self.reminder_thread.start()

        self.timer_thread = threading.Thread(target=self.timer_loop, daemon=True)
        self.timer_thread.start()

        while True:
            try:
                print('Initialising!')
                time.sleep(5)
                self.run()
            except Exception as e:
                mixer.music.load("./sorry.mp3")
                mixer.music.play()

if __name__ == "__main__":
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    response = None
    Voice_cmds = Voice(response, recognizer, microphone)
    Voice_cmds.main()