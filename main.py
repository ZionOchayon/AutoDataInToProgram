# Libraries
import pandas as pd
import pyperclip
import time
import os
import pyautogui

# details
date = "24/2/2010"
xsl_path = r'C:\Users\ziono.DESKTOP-C7FR463\Desktop\elinor.xlsx'
app_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
pc_copy1 = "C:programfiles_"
pc_copy2 = ".xls"
start_point = ""

# import from excel
data = pd.read_excel(xsl_path)

# open the app
os.startfile(app_path)
time.sleep(1)

# start work
pyautogui.click(180, 180)
time.sleep(1)
pyautogui.click(1200, 220)
time.sleep(1)
pyautogui.click(1000, 220)
time.sleep(1)
pyautogui.click(800, 220)
time.sleep(1)
pyautogui.click(200, 180)
time.sleep(1)
pyautogui.doubleClick(810, 540)
time.sleep(1)
pyautogui.click(200, 100)
time.sleep(1)
pyautogui.click(200, 120)
time.sleep(1)
pyautogui.doubleClick(200, 540)
time.sleep(1)

temp = ""
switch = False
for name in data['host name']:
    if type(name) == str and temp != name[0:6]:
        temp = name[0:6]
        if switch == True or temp == start_point:
            switch = True
            pyautogui.rightClick(400, 540)
            time.sleep(1)
            pyautogui.click(800, 500)
            pyperclip.copy(pc_copy1 + temp + "_" + date + pc_copy2)
            pyautogui.click(440, 1350)
            pyautogui.hotkey("ctrlleft", "a")
            pyautogui.hotkey("delete")
            pyautogui.hotkey("ctrlleft", "v")
            pyautogui.click(800, 400)
            time.sleep(1)
            pyautogui.click(100, 180)
            time.sleep(1)
            pyautogui.click(500, 120)
            time.sleep(1)
            pyautogui.click(500, 1000)
            time.sleep(1)
            pyperclip.copy(temp)
            pyautogui.click(50, 100)
            pyautogui.hotkey("ctrlleft", "v")
            pyautogui.hotkey("enter")
            # stop
            pyautogui.click(500, 400)
            time.sleep(5)
