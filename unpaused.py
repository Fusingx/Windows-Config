import time
import pygetwindow

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
        window.resizeTo(0, 200)
        window.moveTo(0, 0)
        break

print('UNPAUSED')
time.sleep(1)