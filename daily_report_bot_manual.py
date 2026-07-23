"""
daily_report_bot.py
Gen AI Architect Program - Assignment 1 (PyAutoGUI Automation)

Everything below happens on screen the way a person would do it - mouse
and keyboard only, no shortcuts that skip the GUI:

  1. Open the Start Menu (Spotlight on Mac) and launch Microsoft Edge.
  2. Type a public weather site's address into the address bar and load it.
  3. Select-all + copy the page text, then pull out the temperature.
  4. Open the Start Menu again and launch Excel.
  5. Type a new row: today's date & time, the fetched weather, a comment.
  6. Save the file as daily_report_YYYY-MM-DD.xlsx through the Save dialog.
  7. Screenshot the finished sheet.
"""

import os
import re
import time
import platform
from datetime import datetime

import pyautogui
import pyperclip

IS_MAC = platform.system() == "Darwin"
MOD = "command" if IS_MAC else "ctrl"
OPEN_SEARCH = ("command", "space") if IS_MAC else ("win",)

CITY = "London"
WEATHER_URL = f"https://wttr.in/{CITY}"

pyautogui.hotkey(*OPEN_SEARCH)
time.sleep(1)
print("Opening Microsoft Edge...")
pyautogui.write("Microsoft Edge", interval=0.05)
time.sleep(1)
pyautogui.press("enter")
time.sleep(4)

pyautogui.hotkey(MOD, "l")            # focus the address bar
time.sleep(0.5)
pyautogui.write(WEATHER_URL, interval=0.02)
pyautogui.press("enter")
time.sleep(3)

print("Copying the page text...")
pyautogui.hotkey(MOD, "a")             # select all page text
time.sleep(0.5)
pyautogui.hotkey(MOD, "c")             # copy it
time.sleep(1)

page_text = pyperclip.paste()
match = re.search(r"[-+]?\d+\s*°C", page_text)
weather = f"{CITY}: {match.group().replace(' ', '')}" if match else "N/A"
print("Weather found:", weather)

# Turn the fetched temperature into a short, meaningful comment
temp_match = re.search(r"[-+]?\d+", weather)
if not temp_match:
    comment = "Weather data unavailable"
else:
    temp = int(temp_match.group())
    if temp >= 25:
        comment = "Warm - great for outdoor activities"
    elif temp >= 10:
        comment = "Mild - a light jacket should do"
    else:
        comment = "Cold - bundle up before heading out"

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
NOW = datetime.now()
TODAY = NOW.strftime("%Y-%m-%d")
DATE_TIME = NOW.strftime("%Y-%m-%d %H:%M:%S")
XLSX_PATH = os.path.join(OUTPUT_DIR, f"daily_report_{TODAY}.xlsx")
SHOT_PATH = os.path.join(OUTPUT_DIR, f"daily_report_{TODAY}.png")

print("Opening Excel...")
pyautogui.hotkey(*OPEN_SEARCH)
time.sleep(1)
pyautogui.write("Excel", interval=0.05)
time.sleep(1)
pyautogui.press("enter")
time.sleep(6)
pyautogui.hotkey("ctrl", "n")          # make sure we land on a blank workbook
time.sleep(2)

print("Typing the report row...")
pyautogui.write("Date & Time", interval=0.03)
pyautogui.press("tab")
pyautogui.write("Fetched Data", interval=0.03)
pyautogui.press("tab")
pyautogui.write("Comment", interval=0.03)
pyautogui.press("enter")               # Excel drops back to column A, next row

pyautogui.write(DATE_TIME, interval=0.03)
pyautogui.press("tab")
# The weather text can contain a "°" symbol PyAutoGUI's write() can't
# type on a plain US keyboard layout, so paste it from the clipboard
# instead of typing it character by character.
pyperclip.copy(weather)
pyautogui.hotkey(MOD, "v")
pyautogui.press("tab")
pyautogui.write(comment, interval=0.03)
pyautogui.press("enter")
time.sleep(1)

print("Saving the workbook...")
pyautogui.press("f12")                 # opens the classic Save-As dialog
time.sleep(2)
pyautogui.write(XLSX_PATH, interval=0.02)
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(2)
pyautogui.press("enter")               # dismiss "Keep current format?" if it shows up
time.sleep(2)
print("Saved workbook:", XLSX_PATH)

print("Taking a screenshot of the sheet...")
img = pyautogui.screenshot()
img.save(SHOT_PATH)
print("Saved screenshot:", SHOT_PATH)

print("Closing the workbook...")
pyautogui.hotkey("ctrl", "w")
time.sleep(2)

print("\nDone! Files created:")
print(" -", XLSX_PATH)
print(" -", SHOT_PATH)