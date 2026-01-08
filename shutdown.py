import win32gui
import win32con
import time
import pygetwindow as gw


print('Starting Shutdown')

def check_windows():
    windows = gw.getAllTitles()
    active_window = win32gui.GetForegroundWindow()
    title = win32gui.GetWindowText(active_window)

    # list of what not to close
    sys_windows = [
    'Settings',
    'Windows Input Experience',
    'Windows Shell Experience Host',
    'Program Manager',
    f'{title}'
    ]

    # filter windows
    filtered_windows = []
    for window in windows:
        if window not in sys_windows and len(window) > 0:
            filtered_windows.append(window)
    
    print('windows checked')
    return filtered_windows

def close_windows(filtered_windows): 
    for window in filtered_windows:
        while True:
            hwnd = win32gui.FindWindow(None, f'{window}') # gets handle of link media window
            if hwnd:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0) # posts msg to handle (hwnd) to close
                print(f'{window} closed')
                break

def force_close_windows(filtered_windows):
    for window in filtered_windows:
        hwnd = win32gui.FindWindow(None, f'{window}')      # gets handle of link media window
        while hwnd:                                            # while the handle exists 
            hwnd = win32gui.FindWindow(None, f'{window}')         # checks to see if the window still exists                
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0) # posts close req to handle
            print(f'{window} closed')
            time.sleep(0.25) 

def main():
    while True:
        filtered_windows = check_windows()
        if filtered_windows:
            close_windows(filtered_windows)
        else:
            break
        force_close_windows(filtered_windows)

main()

