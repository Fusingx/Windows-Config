from pynput import mouse
import pyautogui as p
import win32gui
import win32con
import time

# --- Configuration ---
HOLD_THRESHOLD = 0.1  # Seconds to hold before dragging starts

# --- Global State ---
state = {
    "is_dragging": False,
    "middle_held": False,
    "press_start_time": 0,
    "active_hwnd": None,
    "cursor_offset": (0, 0) # (0,0) means cursor is dead center of window
}

def is_window_maximized(hwnd):
    """Check if a window is maximized"""
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        # placement[1] == 2 is SW_SHOWMAXIMIZED
        return placement[1] == win32con.SW_SHOWMAXIMIZED
    except Exception:
        return False

def get_window_rect(hwnd):
    """Get window coordinates (left, top, right, bottom)"""
    try:
        return win32gui.GetWindowRect(hwnd)
    except:
        return None

def on_click(x, y, button, pressed):
    """Handle mouse clicks"""
    if button == mouse.Button.middle:
        if pressed:
            # --- MOUSE DOWN ---
            state["middle_held"] = True
            state["press_start_time"] = time.time()
        else:
            # --- MOUSE UP ---
            state["middle_held"] = False
            
            if state["is_dragging"]:
                print(f'Dropped at: ({x}, {y})')
                state["is_dragging"] = False
                
                # Snap Logic on Drop
                hwnd = state["active_hwnd"]
                if hwnd:
                    try:
                        if y < 16: # Snap to Top
                            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                        elif y > 1029: # Snap to Bottom (Minimize)   elif y > (p.size()[1] - 20): # Snap to Bottom (Minimize)
                            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    except Exception as e:
                        print(f"Snap error: {e}")
                
                state["active_hwnd"] = None

def on_move(x, y):
    """Handle mouse movement"""
    
    # 1. Check if we are holding the button
    if state["middle_held"]:
        
        # 2. Check if we need to START dragging
        if not state["is_dragging"]:
            if (time.time() - state["press_start_time"]) > HOLD_THRESHOLD:
                try:
                    # Find window under cursor
                    point = (x, y)
                    hwnd = win32gui.WindowFromPoint(point)
                    # Get the root window (in case we clicked a child element)
                    hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)

                    if hwnd:
                        state["active_hwnd"] = hwnd
                        state["is_dragging"] = True
                        
                        # --- CRITICAL LOGIC FOR CENTER SNAP ---
                        if is_window_maximized(hwnd):
                            # 1. Restore window
                            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                            
                            # 2. Wait a tiny bit for Windows to process the restore
                            # (Otherwise we get the 'maximized' size instead of 'normal' size)
                            time.sleep(0.05) 
                            
                            # 3. Get the NEW restored size
                            rect = get_window_rect(hwnd)
                            w = rect[2] - rect[0]
                            h = rect[3] - rect[1]
                            
                            # 4. Force cursor to be dead center
                            state["cursor_offset"] = (0, 0)
                            
                            # 5. Immediately move window to center on cursor
                            new_x = x - (w // 2)
                            new_y = y - (h // 2)
                            win32gui.SetWindowPos(hwnd, None, int(new_x), int(new_y), 0, 0, 
                                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)

                        else:
                            # If window was ALREADY normal, keep the offset
                            # (So the window doesn't jump if you grab it by the corner)
                            rect = get_window_rect(hwnd)
                            win_x = rect[0]
                            win_y = rect[1]
                            w = rect[2] - rect[0]
                            h = rect[3] - rect[1]
                            
                            center_x = win_x + (w // 2)
                            center_y = win_y + (h // 2)
                            
                            # Store the difference between cursor and center
                            state["cursor_offset"] = (x - center_x, y - center_y)

                        print("Drag started")
                        
                except Exception as e:
                    print(f"Start drag error: {e}")

        # 3. PROCESS THE DRAG
        if state["is_dragging"] and state["active_hwnd"]:
            try:
                rect = get_window_rect(state["active_hwnd"])
                if rect:
                    w = rect[2] - rect[0]
                    h = rect[3] - rect[1]
                    
                    # Calculate position: Mouse - HalfWidth - Offset
                    new_x = x - (w // 2) - state["cursor_offset"][0]
                    new_y = y - (h // 2) - state["cursor_offset"][1]

                    win32gui.SetWindowPos(state["active_hwnd"], None, int(new_x), int(new_y), 0, 0, 
                                        win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
            except Exception:
                state["is_dragging"] = False

def run():
    print("Listening... Hold Middle Mouse (0.2s) to drag.")
    with mouse.Listener(on_click=on_click, on_move=on_move) as listener:
        listener.join()

if __name__ == "__main__":
    run()