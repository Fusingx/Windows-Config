import os
import time
import subprocess
import threading
import queue
import win32gui
import win32con
import win32api
import pyautogui as p
import pygetwindow as gw
from pynput import mouse, keyboard

# --- CONFIGURATION ---
CONFIG = {
    "wallpaper_black": os.path.join(os.path.dirname(os.path.realpath(__file__)), "images", "black.png"),
    "wallpaper": os.path.join(os.path.dirname(os.path.realpath(__file__)), "images", "wallpaper.png"),
    "scripts_dir": os.path.dirname(os.path.realpath(__file__)),
    "images_dir": os.path.join(os.path.dirname(os.path.realpath(__file__)), "images"),
    "paths": {
        "vscode": r"C:\Users\Sweetwaters Church\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk",
        "glaze_wm": r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\GlazeWM.lnk"
    },
    "ignore_titles": {'Zen', 'Explorer', 'CapCut', 'Chrome', 'Select exporting path', 'Spotify', 'File Upload'}
}

ALT_COMBOS = {
    'ctrl+q': 'exit',
    'ctrl+m': 'shutdown',
    'ctrl+p': 'pause',
    'ctrl+t': 'taskbar',
    'ctrl+r': 'reload',
    'ctrl+d': 'debug',

    'shift+b': 'browser',
    'shift+m': 'music',
    'shift+f': 'files',
    'shift+v': 'vscode',
    'shift+y': 'yasb',
    'enter': 'terminal',

    'w': 'kill',
    'f': 'fullscreen',
    'r': 'minimize'
}

class AutomationEngine:
    def __init__(self):
        self.command_queue = queue.Queue()
        self.running = True
        self.suppress_listeners = False # Flag to prevent macros from triggering listeners
        
        # State Tracking
        self.modifiers = set()
        self.alt_mode = False
        self.drag_state = {
            "active": False, "middle_held": False, "hwnd": None, 
            "offset": (0, 0), "start_pos": (0, 0)
        }
        
        # Pre-load screen size
        self.screen_w, self.screen_h = p.size()

    # --- LOW LEVEL WINDOW HELPERS ---
    def get_window_at_point(self, x, y):
        try:
            hwnd = win32gui.WindowFromPoint((x, y))
            return win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
        except: return None

    def get_active_window_info(self):
        hwnd = win32gui.GetForegroundWindow()
        windows = gw.getAllTitles()
        return hwnd, win32gui.GetWindowText(hwnd), windows

    def is_maximized(self, hwnd):
        try:
            return win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_SHOWMAXIMIZED
        except: return False

    def safe_find(self, image, timeout=1.0):
        image_path = os.path.join(CONFIG["images_dir"], image)
        start = time.time()
        print(f"Looking for {image}...")
        while time.time() - start < timeout:
            try:
                loc = p.locateOnScreen(image_path, grayscale=True, confidence=0.8) 
                if loc: return loc
            except p.ImageNotFoundException:
                pass
            time.sleep(0.1)
        return None

    # --- LISTENER CALLBACKS ---
    # These run in separate threads. They must be fast and thread-safe.
    
    def on_key_press(self, key):
        if self.suppress_listeners: return

        # Modifier Tracking
        if key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r):
            self.modifiers.add('shift')
        elif key in (keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
            self.modifiers.add('ctrl')
        elif key in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r):
            self.alt_mode = True
            return

        # Alt-Mode Logic
        if self.alt_mode:
            combo_parts = []
            if 'ctrl' in self.modifiers: combo_parts.append('ctrl')
            if 'shift' in self.modifiers: combo_parts.append('shift')
            
            try:
                if hasattr(key, 'vk') and key.vk is not None:
                    char = chr(key.vk).lower() if 65 <= key.vk <= 90 else key.char.lower()
                    combo_parts.append(char)
                else:
                    combo_parts.append(key.char.lower())
            except AttributeError:
                clean_key = str(key).replace('Key.', '')
                combo_parts.append(clean_key)

            combo_str = "+".join(combo_parts)
            
            if combo_str in ALT_COMBOS:
                self.command_queue.put(("cmd", ALT_COMBOS[combo_str]))

    def on_key_release(self, key):
        if key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r):
            self.modifiers.discard('shift')
        if key in (keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
            self.modifiers.discard('ctrl')
        if key in (keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r):
            self.alt_mode = False

    def on_click(self, x, y, button, pressed):
        if self.suppress_listeners: return

        if button == mouse.Button.middle:
            if pressed:
                self.drag_state["middle_held"] = True
                self.drag_state["start_pos"] = (x, y)
            else:
                self.drag_state["middle_held"] = False
                if self.drag_state["active"]:
                    # End Drag / Snap Logic
                    self.command_queue.put(("snap_check", (self.drag_state["hwnd"], y)))
                    self.drag_state["active"] = False
                    self.drag_state["hwnd"] = None
        
        elif pressed:
            if button == mouse.Button.x1:
                self.command_queue.put(("mouse_macro", "x1"))
            elif button == mouse.Button.x2:
                self.command_queue.put(("mouse_macro", "x2"))

    def on_move(self, x, y):
        # High frequency event - keep logic minimal
        ds = self.drag_state
        if ds["middle_held"]:
            if not ds["active"]:
                # Check threshold
                dx = abs(x - ds["start_pos"][0])
                dy = abs(y - ds["start_pos"][1])
                if dx > 10 or dy > 10:
                    hwnd = self.get_window_at_point(x, y)
                    if hwnd:
                        ds["active"] = True
                        ds["hwnd"] = hwnd
                        
                        # Handle maximized windows before dragging
                        if self.is_maximized(hwnd):
                            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                            time.sleep(0.05) # Allow OS to repaint
                            # Recalculate center offset
                            rect = win32gui.GetWindowRect(hwnd)
                            w, h = rect[2] - rect[0], rect[3] - rect[1]
                            ds["offset"] = (w // 2, h // 2)
                            # Jump window to cursor
                            win32gui.SetWindowPos(hwnd, None, x - (w//2), y - (h//2), 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
                        else:
                            rect = win32gui.GetWindowRect(hwnd)
                            ds["offset"] = (x - rect[0], y - rect[1])
            
            # Perform Drag
            if ds["active"] and ds["hwnd"]:
                try:
                    rect = win32gui.GetWindowRect(ds["hwnd"])
                    w, h = rect[2] - rect[0], rect[3] - rect[1]
                    new_x = x - ds["offset"][0]
                    new_y = y - ds["offset"][1]
                    win32gui.SetWindowPos(ds["hwnd"], None, int(new_x), int(new_y), 0, 0, 
                                        win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
                except Exception:
                    ds["active"] = False

    def on_scroll(self, x, y, dx, dy):
        if self.alt_mode:
            if dy > 0: self.command_queue.put(("cmd", "alt_scroll_up"))
            elif dy < 0: self.command_queue.put(("cmd", "alt_scroll_down"))

    # --- ACTION HANDLERS (MAIN THREAD) ---

    def exec_snap(self, hwnd, y):
        try:
            if y < 16: # Top Snap
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
            elif y > (self.screen_h - 20): # Bottom Snap
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        except Exception as e:
            print(f"Snap failed: {e}")

    def exec_taskbar_toggle(self):
        # Moves cursor, clicks taskbar settings. 
        # Note: Direct registry modification is cleaner but requires explorer restart.
        # Keeping UI automation as requested.
        print('Toggling Taskbar')
        current_pos = p.position()
        p.moveTo(1317, 1079)
        time.sleep(0.2)
        p.rightClick()
        p.doubleClick(1340, 1011)
        
        behaviors = self.safe_find('behaviors.jpg', 2)
        if behaviors:
            p.moveTo(behaviors)
            time.sleep(0.2)
            p.click(behaviors)
            hide_opt = self.safe_find('hide.jpg', 1)
            if hide_opt: p.click(hide_opt)
            p.hotkey('alt', 'f4')
        p.moveTo(current_pos)

    def exec_kill_all(self):
        print("Closing user apps...")
        sys_apps = ['Settings', 'Windows Input Experience', 'Program Manager', 'SearchHost']
        
        def close_handles():
            top_level_windows = []
            win32gui.EnumWindows(lambda hwnd, list: list.append((hwnd, win32gui.GetWindowText(hwnd))), top_level_windows)
            
            for hwnd, title in top_level_windows:
                if win32gui.IsWindowVisible(hwnd) and title and title not in sys_apps:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    print(f"Closed: {title}")
        
        close_handles()
        # Fallback force close pass could go here

    def exec_shutdown_routine(self):
        hwnd, title, windows = self.get_active_window_info()
        print("Shutting down sequence...")
        # Close Yasb/Glaze
        if 'YasbBar' in windows:
            os.system("yasbc stop")
            os.system("glazewm command wm-exit")
            self.exec_taskbar_toggle()
            while 'YasbBar' in windows:
                hwnd, title, windows = self.get_active_window_info()
                print('waiting for yasb to close')
                time.sleep(0.1)
        
        # Hide Desktop Icons
        p.moveTo(self.screen_w - 5, self.screen_h - 1)
        p.click(self.screen_w - 5, self.screen_h - 1) # Show desktop corner
        time.sleep(0.5)
        p.rightClick(self.screen_w // 2, self.screen_h // 2)
        p.press(['right', 'up', 'enter'])
        
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, CONFIG["wallpaper_black"], 1+2)
        self.exec_kill_all()
        os.system('exit')
        exit()
    
    # --- APP MACROS ---
    def run_mouse_macro(self, button):
        hwnd, title, windows = self.get_active_window_info()
        print(f"Macro {button} on: {title}")
        
        if any(ign in title for ign in CONFIG["ignore_titles"]):
            return

        # Prevent recursive listener triggering during macro execution
        self.suppress_listeners = True 
        
        try:
            if 'PowerPoint' in title:
                if button == 'x1':
                    p.click()
                    mouse_pos = p.position()
                    if loc := self.safe_find('colour.jpg'):
                        p.click(loc)
                        p.moveTo(mouse_pos)
                        p.hotkey('ctrl', 'b') 
                        p.hotkey('ctrl', 'u')
                else: # x2
                    p.click()
                    p.hotkey('ctrl', 'a')
                    mouse_pos = p.position()
                    if loc := self.safe_find('paste.jpg'): p.click(loc)
                    p.moveTo(mouse_pos)

            elif 'ProPresenter' in title:
                if button == 'x1':
                    p.rightClick()
                    mouse_pos = p.position()
                    if loc := self.safe_find('edit.jpg'):
                        p.click(loc)
                        p.moveTo(mouse_pos)
                else: # x2
                    p.rightClick()
                    mouse_pos = p.position()
                    if loc := self.safe_find('edit.jpg'):
                        p.moveTo(loc)
                        time.sleep(0.1)
                        p.click(loc)
                        p.hotkey('ctrl', 'a')
                        p.hotkey('alt', 'v')
                    p.press('esc', presses=2)

            elif 'Lightroom' in title:
                if button == 'x1': p.press('r')
                else:
                    orig_pos = p.position()
                    if loc := self.safe_find('notcrop.jpg'): p.click(loc)
                    p.press('r')
                    if loc := self.safe_find('original.jpg'): 
                        p.click(loc)
                        if loc2 := self.safe_find('1x1.jpg'): p.click(loc2)
                    p.moveTo(orig_pos)
            
            else: # Default Media Control
                if button == 'x1': p.press('playpause')
                else: p.press('nexttrack')
                
        finally:
            self.suppress_listeners = False

    # --- MAIN LOOP ---
    def run(self):
        # Set Wallpaper
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, CONFIG["wallpaper"], 1+2)
        print("Engine Started. Listening...")

        # Start Threads
        m_listener = mouse.Listener(on_click=self.on_click, on_move=self.on_move, on_scroll=self.on_scroll)
        k_listener = keyboard.Listener(on_press=self.on_key_press, on_release=self.on_key_release)
        m_listener.start()
        k_listener.start()

        # Command Consumer Loop
        while self.running:
            try:
                # Blocks until an item is available - 0% CPU usage while waiting
                msg_type, data = self.command_queue.get(timeout=1) 
                
                if msg_type == "cmd":
                    self.handle_global_command(data)
                elif msg_type == "snap_check":
                    self.exec_snap(data[0], data[1])
                elif msg_type == "mouse_macro":
                    self.run_mouse_macro(data)
                
                self.command_queue.task_done()
            
            except queue.Empty:
                continue
            except KeyboardInterrupt:
                break
        
        m_listener.stop()
        k_listener.stop()

    def handle_global_command(self, cmd):
        print(f"Executing: {cmd}")
        if cmd == 'exit':
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, CONFIG["wallpaper_black"], 1+2)
            self.running = False
        elif cmd == 'pause':
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, CONFIG["wallpaper_black"], 1+2)
            subprocess.Popen(["pythonw", "paused.py"], cwd=CONFIG["scripts_dir"])
            self.running = False
        elif cmd == 'shutdown': self.exec_shutdown_routine()
        elif cmd == 'reload': 
            os.system(f'start "" pyw "{os.path.join(CONFIG["scripts_dir"], "main.py")}"')
            self.running = False
        elif cmd == 'debug':
            os.system(f'start "" py "{os.path.join(CONFIG["scripts_dir"], "main.py")}"')
            self.running = False
        elif cmd == 'taskbar': self.exec_taskbar_toggle()
        elif cmd == 'terminal': os.startfile('PowerShell')
        elif cmd == 'browser': os.startfile('zen')
        elif cmd == 'music': os.startfile('spotify')
        elif cmd == 'files': os.startfile('explorer')
        elif cmd == 'vscode': os.startfile(CONFIG["paths"]["vscode"])
        elif cmd == 'yasb':
            windows = gw.getAllTitles()
            if 'YasbBar' in windows:
                print('Glaze / YASB Closing')
                os.system("yasbc stop")
                os.system("glazewm command wm-exit")
                self.exec_taskbar_toggle
            else:
                print('Glaze / YASB Starting')
                self.exec_taskbar_toggle
                os.startfile('yasb')
                os.startfile(CONFIG["paths"]["glaze_wm"])
                
        elif cmd == 'kill':
            hwnd = win32gui.GetForegroundWindow()
            if hwnd: win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        elif cmd == 'fullscreen':
            hwnd = win32gui.GetForegroundWindow()
            mode = win32gui.GetWindowPlacement(hwnd)[1]
            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL if mode == win32con.SW_SHOWMAXIMIZED else win32con.SW_MAXIMIZE)
        elif cmd == 'minimize':
            win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE)
        elif cmd == 'alt_scroll_up': p.hotkey('alt', 's')
        elif cmd == 'alt_scroll_down': p.hotkey('alt', 'a')

if __name__ == "__main__":
    app = AutomationEngine()
    app.run()