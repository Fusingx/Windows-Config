import os
import win32gui
import win32con
import pyautogui as p
import time
import random
import pygetwindow

def find(element, timeout=0):
    if timeout != 0:
        TIMEOUT = timeout
        START_TIME = time.time()
    while True:
        try:
            coord = p.locateOnScreen(f'{element}', confidence=0.8)
            if coord:
                return coord
        except:
            p.ImageNotFoundException
        
        if timeout != 0:
            if time.time() - START_TIME > TIMEOUT:
                print(f'Failed to locate "{element}"')
                p.press('right')
                return False
            
print('Welcome')

p.hotkey('win', 'ctrl', 't') # make the terminal always on top using powertoys

window = '' # get a handle of the terminal window to move and resize it
while True:
    try:
        window = pygetwindow.getWindowsWithTitle("Alacritty")[0]
    except:
        pass

    try:
        window = pygetwindow.getWindowsWithTitle(r"C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe")[0]
    except:
        pass

    try:
        window = pygetwindow.getWindowsWithTitle(r'C:\WINDOWS\system32\cmd.exe')[0]
    except:
        pass

    if window:
        window.resizeTo(500, 500)
        window.moveTo(0, 0)
        break

p.rightClick(x=1823, y=66) # hide desktop icons
#p.click(x=1704, y=86)
p.click(find('view.jpg'))
time.sleep(0.4)
#p.click(x=1470, y=258)
p.click(find('icons.jpg'))

os.startfile('zen.exe') # open zen and spotify
os.startfile('spotify.exe')

for i in range(100): # welcome text on terminal
    colour = random.randrange(31, 39)
    print(f'\033[{colour}mWelcome !\033[0m')
    time.sleep(0.02)

find('spotify.jpg') # makes sure spotify & zen are actually open before trying to minimize them
win32gui.ShowWindow(win32gui.FindWindow(None, r'Zen Browser'), win32con.SW_MINIMIZE) # minimize zen & spotify
win32gui.ShowWindow(win32gui.FindWindow(None, r'Spotify Premium'), win32con.SW_MINIMIZE)

while True: # close terminal
    hwnd = win32gui.FindWindow(None, r'C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe')
    if hwnd:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    hwnd = win32gui.FindWindow(None, r'C:\WINDOWS\system32\cmd.exe')
    if hwnd:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(0.05)
