import os
import win32gui
import win32con
import time
import subprocess
import pygetwindow as gw
from pynput import mouse
from pynput import keyboard
import pyautogui as p

# --- 1. DRAG CONFIGURATION & GLOBALS ---
HOLD_THRESHOLD = 0
DRAG_THRESHOLD = 10
drag_state = {
    "is_dragging": False,
    "middle_held": False,
    "press_start_time": 0,
    "active_hwnd": None,
    "cursor_offset": (0, 0),
    "start_x": 0,
    "start_y": 0
}

# --- 2. EXISTING GLOBALS ---
windows = gw.getAllTitles()
active_window = win32gui.GetForegroundWindow()
title = win32gui.GetWindowText(active_window)
alt_mode = False
current_modifiers = set()
button_signal = None
listener_running = True
paused = False

alt_combos = {
    'f': 'fullscreen',
    'shift+f': 'files',
    'shift+v': 'vscode',
    'w': 'kill',
    'enter': 'terminal',
    'shift+r': 'minimize',
    'shift+b': 'browser',
    'shift+m': 'music'
}

# --- 3. HELPER FUNCTIONS ---

def is_window_maximized(hwnd):
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        return placement[1] == win32con.SW_SHOWMAXIMIZED
    except:
        return False

def get_window_rect(hwnd):
    try:
        return win32gui.GetWindowRect(hwnd)
    except:
        return None

# --- 4. FIND FUNCTION ---

def find(element, timeout=0):
    if timeout != 0:
        TIMEOUT = timeout
        START_TIME = time.time()
    print(f'Attempting to locate "{element}"')
    while True:
        try:
            coord = p.locateOnScreen(f'{element}', confidence=0.8)
            if coord:
                print(f'"{element}" Located')
                return coord
        except:
            pass # p.ImageNotFoundException
        
        if timeout != 0:
            if time.time() - START_TIME > TIMEOUT:
                print(f'Failed to locate "{element}"')
                p.press('right')
                return False

# --- 5. PYNPUT LISTENERS ---

def on_click(x, y, button, pressed):
    global button_signal, listener_running, drag_state
    #ignore = ('Zen', 'Chrome')
    
    # --- MIDDLE MOUSE LOGIC (Does not stop listener) ---
    if button == mouse.Button.middle:
        #title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        #if any(ignore_string in title for ignore_string in ignore):
            #return
        if pressed:
            drag_state["middle_held"] = True
            drag_state["press_start_time"] = time.time()
            drag_state["start_x"] = x
            drag_state["start_y"] = y
        else:
            drag_state["middle_held"] = False
            if drag_state["is_dragging"]:
                print(f'Dropped at: ({x}, {y})')
                drag_state["is_dragging"] = False
                
                # Snap Logic on Drop
                hwnd = drag_state["active_hwnd"]
                if hwnd:
                    try:
                        if y < 16: # Snap to Top
                            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                            win32gui.BringWindowToTop(hwnd)
                            win32gui.SetForegroundWindow(hwnd)
                        elif y > 1029: # Snap to Bottom         elif y > (p.size()[1] - 20): 
                            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    except Exception as e:
                        print(f"Snap error: {e}")
                drag_state["active_hwnd"] = None

    # --- X1 / X2 LOGIC (Stops listener to run macros) ---
    if pressed:
        if button == mouse.Button.x1:  # Mouse 4
            print(f"Mouse Button 4 clicked at ({x}, {y})")
            button_signal = "x1"
            listener_running = False
            return False  # stop listener
        elif button == mouse.Button.x2:  # Mouse 5
            print(f"Mouse Button 5 clicked at ({x}, {y})")
            button_signal = "x2"
            listener_running = False
            return False  # stop listener

def on_move(x, y):
    global drag_state
    
    # 1. Check if holding middle button
    if drag_state["middle_held"]:
        delta_x = abs(x - drag_state["start_x"])
        delta_y = abs(y - drag_state["start_y"])
        
        
        # 2. Start Dragging if threshold met
        if not drag_state["is_dragging"]:
            #if (time.time() - drag_state["press_start_time"]) > HOLD_THRESHOLD:
            if delta_x > DRAG_THRESHOLD or delta_y > DRAG_THRESHOLD:
                try:
                    point = (x, y)
                    hwnd = win32gui.WindowFromPoint(point)
                    hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)

                    if hwnd:
                        drag_state["active_hwnd"] = hwnd
                        drag_state["is_dragging"] = True
                        
                        if is_window_maximized(hwnd):
                            # Restore and Center Snap
                            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                            time.sleep(0.05)
                            rect = get_window_rect(hwnd)
                            w = rect[2] - rect[0]
                            h = rect[3] - rect[1]
                            drag_state["cursor_offset"] = (0, 0)
                            
                            new_x = x - (w // 2)
                            new_y = y - (h // 2)
                            win32gui.SetWindowPos(hwnd, None, int(new_x), int(new_y), 0, 0, 
                                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
                        else:
                            # Keep relative offset
                            rect = get_window_rect(hwnd)
                            w = rect[2] - rect[0]
                            h = rect[3] - rect[1]
                            center_x = rect[0] + (w // 2)
                            center_y = rect[1] + (h // 2)
                            drag_state["cursor_offset"] = (x - center_x, y - center_y)

                        print("Drag started")
                except Exception as e:
                    print(f"Start drag error: {e}")

        # 3. Update Window Position
        if drag_state["is_dragging"] and drag_state["active_hwnd"]:
            try:
                rect = get_window_rect(drag_state["active_hwnd"])
                if rect:
                    w = rect[2] - rect[0]
                    h = rect[3] - rect[1]
                    new_x = x - (w // 2) - drag_state["cursor_offset"][0]
                    new_y = y - (h // 2) - drag_state["cursor_offset"][1]

                    win32gui.SetWindowPos(drag_state["active_hwnd"], None, int(new_x), int(new_y), 0, 0, 
                                        win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
            except Exception:
                drag_state["is_dragging"] = False

# --- Keyboard handlers ---
def on_key_press(key):
    global alt_mode, button_signal, listener_running

    # Track Shift modifier
    if key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r):
        current_modifiers.add('shift')

    # Check if Alt pressed → enter Alt mode
    elif key in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r):
        alt_mode = True
        return

    # Check Alt-mode combos
    elif alt_mode:
        combo = ''
        if 'shift' in current_modifiers:
            combo += 'shift+'

        try:
            combo += key.char.lower()
        except AttributeError:
            if key == keyboard.Key.enter:
                combo += 'enter'
            else:
                combo += str(key).replace('Key.', '')

        if combo in alt_combos:
            button_signal = alt_combos[combo]
            listener_running = False
            print(f"Alt combo triggered: {button_signal}")

def on_key_release(key):
    global alt_mode

    # Remove Shift from modifiers
    if key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r):
        current_modifiers.discard('shift')

    # Exit Alt mode when Alt released
    if key in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r):
        alt_mode = False

# --- Listener starter ---
def startListener():
    global button_signal, listener_running

    button_signal = None
    listener_running = True

    # Start mouse listener
    mouse_listener = mouse.Listener(on_click=on_click, on_move=on_move)
    mouse_listener.start()

    # Start keyboard listener
    keyboard_listener = keyboard.Listener(
        on_press=on_key_press,
        on_release=on_key_release
    )
    keyboard_listener.start()

    print("Listening (Alt-mode)...")
    while listener_running:
        time.sleep(0.01)

    mouse_listener.stop()
    keyboard_listener.stop()
    mouse_listener.join()
    keyboard_listener.join()

    print(f"Listener stopped, returning signal: {button_signal}")
    return button_signal

# --- 6. APP SCRIPTS ---

# powerpoint
def x1_powerpoint():
    p.click()
    mouse_pos = p.position()
    coord = find('colour.jpg', 2)
    if not coord:
        return
    p.click(coord)
    p.moveTo(mouse_pos)
    p.hotkey('ctrl', 'b')
    p.hotkey('ctrl', 'u')
    main()

def x2_powerpoint():
    p.click()
    p.hotkey('ctrl', 'a')
    mouse_pos = p.position()
    coord = find('paste.jpg', 2)
    if not coord:
        return
    p.click(coord)
    p.moveTo(mouse_pos)
    main()

# propresenter
def x1_propresenter():
    p.rightClick()
    mouse_pos = p.position()
    p.click(find('edit.jpg'))
    p.moveTo(mouse_pos)
    # p.click()
    # p.hotkey('ctrl', 'a')
    # p.hotkey('alt', 'v')
    main()

def x2_propresenter():
    p.rightClick()
    mouse_pos = p.position()
    p.click(find('edit.jpg'))
    p.moveTo(mouse_pos)
    time.sleep(0.1) # ---------
    p.click()
    p.hotkey('ctrl', 'a')
    p.hotkey('alt', 'v')
    p.press('esc')
    p.press('esc')
    main()

# lightroom
def x1_lightroom():
    p.press('r')

def x2_lightroom():
    pos = p.position()
    notcrop = find('notcrop.jpg')
    p.moveTo(notcrop)
    p.click(notcrop)
    p.press('r')
    original = find('original.jpg', timeout=1)
    if original == False:
        p.moveTo(pos)
        return
    p.click(original)
    p.click(find('1x1.jpg'))
    p.moveTo(pos)

# spotify
def x1_spotify():
    p.press('playpause')
    main()

def x2_spotify():
    p.press('nexttrack')
    main()

# hide taskbar
def taskbar():
    p.moveTo(x=1317, y=1079)
    time.sleep(0.5)
    p.rightClick(x=1317, y=1079) 
    p.doubleClick(x=1340, y=1011)
    behaviors = find('behaviors.jpg')
    p.moveTo(behaviors)
    time.sleep(0.5)
    p.click(behaviors)
    p.click(find('hide.jpg'))
    p.hotkey('alt', 'f4')

# kill
def kill():
    #active_window = win32gui.GetForegroundWindow()
    #hwnd = win32gui.FindWindow(None, f'{active_window}') # gets handle of link media window
    #if hwnd:
    #    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0) # posts msg to handle (hwnd) to close
    #    print(f'{active_window} closed')
    p.hotkey('alt', 'f4')

# fullscreen
def fullscreen():
    hwnd = win32gui.GetForegroundWindow()
    if is_window_maximized(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
    else:
        win32gui.ShowWindow(hwnd, win32con.SHOW_FULLSCREEN)

# minimize
def minimize():
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    return hwnd

# kill all
def kill_all():
    print('Terminating Windows')

    def check_windows():
        windows = gw.getAllTitles()
        active_window = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(active_window)

        # list of what not to close
        sys_windows = [
        'Settings',
        'Windows Input Experience',
        'Windows Shell Experience Host',
        'Program Manager'
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
            while hwnd:                                    # while the handle exists 
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

# --- 7. MAIN LOOP ---

def main():
    ignore = ['Zen', 'Explorer', 'CapCut', 'Chrome', 'Select exporting path', 'Spotify']
    global paused
    while True:
        signal = startListener()
        
        if signal == 'pause': # alt ctrl p
            if paused == False:
                win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, r"C:\Users\Sweetwaters Church\Pictures\wallpaper.png", 1+2)
                os.system('py paused.py')
                paused = True
            else:
                win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, r"C:\Users\Sweetwaters Church\Pictures\wallpaper2.png", 1+2)
                os.system('py unpaused.py')
                paused = False

        elif paused == True:
            continue 
        
        elif signal == "exit": # alt ctrl q
            os.system('exit')
            print('Exitting')
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, r"C:\Users\Sweetwaters Church\Pictures\wallpaper.png", 1+2)
            print('Wallpaper Set')
            exit()

        elif signal == "shutdown": # alt ctrl m
            print('Shutting Down')
            if 'GlazeWM' in windows:
                os.startfile(r"C:\Users\Sweetwaters Church\Jordan\Scripts\WM-Exit\WM-Exit.cmd")
                while 'GlazeWM' in windows:
                    time.sleep(0.1)
            else:
                p.click(1914, 1054)
                time.sleep(1)
                p.rightClick(x=1823, y=66) 
                p.click(find('view.jpg'))
                time.sleep(0.4)
                p.click(find('icons.jpg'))
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, r"C:\Users\Sweetwaters Church\Pictures\wallpaper.png", 1+2)
            print('Wallpaper Set')
            kill_all()
            exit()

        elif signal == 'reload': # alt ctrl r
            print('reloading config')
            os.system(r'start "" pyw "C:\Users\Sweetwaters Church\Jordan\Scripts\Shortcuts\main.py"')
            exit()

        elif signal == 'taskbar': # alt ctrl t
            taskbar()

        elif signal == 'terminal': # alt enter
            try:
                hwnd = win32gui.FindWindow(None, "Windows Input Experience")
                if hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # restore if minimized
            except:
                pass
            os.startfile('PowerShell')
            

        elif signal == 'browser': # alt shift b
            os.startfile('zen')

        elif signal == 'music': # alt shift m
            os.startfile('spotify')

        elif signal == 'files': # alt shift f
            os.startfile('explorer')

        elif signal == 'vscode': # alt shift v
            os.startfile(r"C:\Users\Sweetwaters Church\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk")

        elif signal == 'kill': # alt w
            kill()

        elif signal == 'fullscreen': # alt f
            fullscreen()

        elif signal == 'minimize': # alt r
            minimize()
            
        active_window = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(active_window)
        
        if signal == "x1":
            print(f"Mouse 4 action triggered for {title}.")
            if any(ignore_string in title for ignore_string in ignore):
                continue
            if 'PowerPoint' in title:
                x1_powerpoint()
            elif 'ProPresenter' in title:
                x1_propresenter()
            elif 'Lightroom' in title:
                x1_lightroom()
            else:
                x1_spotify()

        elif signal == "x2":
            print("Mouse 5 action triggered.")
            if any(ignore_string in title for ignore_string in ignore):
                continue
            if 'PowerPoint' in title:
                x2_powerpoint()
            elif 'ProPresenter' in title:
                x2_propresenter()
            elif 'Lightroom' in title:
                x2_lightroom()
            else:
                x2_spotify()

if __name__ == "__main__":
    win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, r"C:\Users\Sweetwaters Church\Pictures\wallpaper2.png", 1+2)
    print('Wallpaper Set')
    main()