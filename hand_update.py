import speech_recognition as sr
import os
import pyautogui
import glob
import webbrowser
import pygetwindow as gw
import psutil 
import subprocess
import time
import ctypes
import datetime
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Function to decrease the system volume
def decrease_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = max(0, current_volume - 0.1)  # Adjust as needed
    volume.SetMasterVolumeLevelScalar(new_volume, None)

# Function to increase the system volume
def increase_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = min(1, current_volume + 0.1)  # Adjust as needed
    volume.SetMasterVolumeLevelScalar(new_volume, None)

def sleep_system():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def lock_screen():
    ctypes.windll.user32.LockWorkStation()

def restart_system():
    os.system("shutdown /r /t 1")

# Function to execute commands based on recognized speech
def execute_command(command):
    if "sleep system" in command:
        sleep_system()
    elif "restart system" in command:
        restart_system()
    elif "lock screen" in command:
        lock_screen()
    elif "maximize window" in command or "maximise window" in command: 
        maximize_window()
    elif "minimize window" in command or "minimise window" in command:  
        minimize_window()
    elif "switch window " in command:
        switch_window()
    elif "take screenshot" in command:
        take_screenshot()
    elif "increase volume" in command:
        increase_volume()
    elif "decrease volume" in command:
        decrease_volume()
    elif "next slide" in command:
        next_slide()
    elif "previous slide" in command:
        previous_slide()
    elif "close window" in command:
        close_window()
    elif "mute volume" in command:
        mute_volume()
    elif "unmute volume" in command:
        unmute_volume()
    elif "set timer" in command:
        set_timer()
    elif "shut down" in command or "shutdown" in command:  # Update condition for shutting down
        shut_down()

def close_window():
    time.sleep(1)
    active_window = gw.getActiveWindow()
    if active_window is not None:
        active_window.close()
    else:
        print("No active window found.")


def next_slide():
    pyautogui.press('right')
    time.sleep(1)

def previous_slide():
    pyautogui.press('left')
    time.sleep(1)


# Function to maximize the currently focused window
def maximize_window():
    active_window = gw.getActiveWindow()
    if active_window:
       active_window.maximize()

# Function to minimize the currently focused window
def minimize_window():
    active_window = gw.getActiveWindow()
    if active_window:
       active_window.minimize()

# Function to take a screenshot
def take_screenshot():
    try:
        # Specify the directory where you want to save the screenshot
        screenshot_directory = r"C:\Users\vinay bodem\OneDrive\Pictures\Screenshots"
        
        # Ensure the directory exists, if not, create it
        if not os.path.exists(screenshot_directory):
            os.makedirs(screenshot_directory)

        # Generate a unique filename using current timestamp
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        screenshot_filename = f"screenshot_{timestamp}.png"

        # Construct the full path to save the screenshot
        screenshot_path = os.path.join(screenshot_directory, screenshot_filename)

        # Take the screenshot and save it with the unique filename
        pyautogui.screenshot(screenshot_path)
        print("Screenshot saved successfully:", screenshot_filename)

    except Exception as e:
        print("Error taking screenshot:", e)


# Function to switch to a specified window
def switch_window():
    pyautogui.hotkey("alt", "tab")
    time.sleep(1)# Add a delay to ensure the window switch is successful


# Function to mute the system volume
def mute_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(True, None)

def unmute_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(False, None)


# Function to play music
def play_music():
    # Add logic to play music from a specified source
    pass

# Function to pause music
def pause_music():
    # Add logic to pause currently playing music
    pass

# Function to play the next track
def next_track():
    # Add logic to play the next track in the playlist
    pass

# Function to play the previous track
def previous_track():
    # Add logic to play the previous track in the playlist
    pass

# Function to set a timer
import time

# Function to set a timer
def set_timer():
    time.sleep(15)
    print("Timer expired!")

# Function to shut down the system
def shut_down():
    os.system("shutdown /s /t 1")


import cv2
import time
import os
import HandTrackingModule as htm

wCam, hCam = 1280, 720

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

detector = htm.HandDetector(detectionCon=0.75)
tipIds = [4, 8, 12, 16, 20]
fingerNames = ['Thumb', 'Index', 'Middle', 'Ring', 'Pinky']

while True:
    try:
        success, img = cap.read()
        img = detector.findHands(img)
        lmList = detector.findPosition(img, draw=False)

        if len(lmList) != 0:
            fingers = detector.fingersUp()
            totalFingers = fingers.count(1)

            # Map finger gestures to functions
            if fingers == [1, 0, 0, 0, 0]:  # Thumb up
                increase_volume()
            elif fingers == [0, 1, 1, 1, 1]:  # All fingers except Thumb up
                decrease_volume()
            elif fingers == [0, 1, 0, 0, 0]:  # Index finger up
                switch_window()
            elif fingers == [1, 1, 0, 0, 0]:  # Thumb and Index fingers up
                take_screenshot()
            elif fingers == [1, 1, 1, 1, 1]:  # All fingers up
                shut_down()
            elif fingers == [1, 1, 1, 0, 0]:  # Thumb, Index, Middle fingers up
                mute_volume()
            elif fingers == [0, 0, 0, 0, 1]:  # Pinky finger up
                unmute_volume()
            elif fingers == [0, 1, 0, 1, 1]:  # Index and Ring fingers up
                next_slide()
            elif fingers == [0, 1, 1, 0, 0]:  # Index and Middle fingers up
                previous_slide()
            elif fingers == [1, 1, 0, 1, 0]:  # Thumb, Index, and Ring fingers up
                maximize_window()
            elif fingers == [0, 0, 1, 0, 0]:  # Middle finger up
                minimize_window()
            elif fingers == [0, 0, 0, 1, 1]:  # Ring and Pinky fingers up
                lock_screen()
            elif fingers == [1, 0, 1, 0, 1]:  # Thumb and Middle fingers up
                restart_system()
            elif fingers == [1, 0, 1, 1, 0]:  # Thumb, Middle, and Ring fingers up
                set_timer()
            elif fingers == [1, 0, 0, 1, 1]:  # Thumb, Ring, and Pinky fingers up
                close_window()


            for id, finger in enumerate(fingers):
                if finger == 1:
                    cv2.putText(img, fingerNames[id], (45, 375 + id * 50), cv2.FONT_HERSHEY_PLAIN,
                                3, (255, 0, 0), 3)

            cv2.putText(img, f"Total Fingers: {totalFingers}", (45, 100), cv2.FONT_HERSHEY_PLAIN,
                        3, (0, 255, 0), 3)

        cv2.imshow("Image", img)
        cv2.waitKey(1)

    except Exception as e:
        print(f"Error: {e}")
        break  # Exit the loop if an error occurs

cap.release()
cv2.destroyAllWindows()

