import os
import win32gui
import win32con
import pyautogui as p
import time
import random
import pygetwindow




while True:
    win32gui.ShowWindow(win32gui.FindWindow(None, r'Zen Browser'), win32con.SW_MINIMIZE) # minimize zen & spotify
    win32gui.ShowWindow(win32gui.FindWindow(None, r'Spotify Premium'), win32con.SW_MINIMIZE)
    time.sleep(0.02)

input('end')