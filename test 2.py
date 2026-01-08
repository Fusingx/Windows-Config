from pynput import mouse
import pyautogui as p
import win32gui
import win32con
import time

# Global variables for dragging
is_dragging = False
window_initial_pos = None
cursor_initial_offset = None
active_window = None
listener_running = False

def is_window_maximized(hwnd):
    """Check if a window is maximized"""
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        # placement[1] is the showCmd which indicates window state
        return placement[1] == win32con.SW_SHOWMAXIMIZED
    except Exception as e:
        print(f"Error checking window state: {e}")
        return False
    
def on_click(x, y, button, pressed):
    global is_dragging, window_initial_pos, cursor_initial_offset, active_window, listener_running
    window = win32gui.GetForegroundWindow()
    window_text = win32gui.GetWindowText(window)

    if button == mouse.Button.middle:
        if pressed: 
                try:
                    active_window = p.getActiveWindow()
                    if active_window:
                        was_maximized = is_window_maximized(window)
                        
                        if was_maximized or win32gui.IsIconic(window):  # Check if maximized or minimized
                            while True:
                                hwnd = win32gui.FindWindow(None, f'{window_text}')
                                if hwnd:
                                    win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)  # unmaximizes/restores
                                    break
                                time.sleep(0.1)
                            
                            # Wait a bit for window to restore
                            time.sleep(0.05)
                            
                            # Get the updated window info after restore
                            active_window = p.getActiveWindow()
                            if active_window:
                                # Center the window at cursor position
                                new_x = x - active_window.width // 2
                                new_y = y - active_window.height // 2
                                active_window.moveTo(new_x, new_y)
                                
                                # Update tracking variables for smooth dragging
                                window_initial_pos = (new_x, new_y)
                                window_center_x = new_x + active_window.width // 2
                                window_center_y = new_y + active_window.height // 2
                                cursor_initial_offset = (x - window_center_x, y - window_center_y)
                        
                        else:
                            # Window was already normal, set up tracking normally
                            window_initial_pos = (active_window.left, active_window.top)
                            window_center_x = active_window.left + active_window.width // 2
                            window_center_y = active_window.top + active_window.height // 2
                            cursor_initial_offset = (x - window_center_x, y - window_center_y)
                        
                        is_dragging = True
                        print(f'Started dragging window from: ({x}, {y})')
                        
                except Exception as e:
                    print(f"Error: {e}")
        else:
            # Stop dragging
            is_dragging = False
            print(f'Stopped dragging window at: ({x}, {y})') 
            
            # Auto-maximize when dragged to top of screen
            if y < 16:
                while True:
                    hwnd = win32gui.FindWindow(None, f'{window_text}')
                    if hwnd:
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # maximizes
                        break
                    time.sleep(0.1)
            elif y > 1029:
                while True:
                    hwnd = win32gui.FindWindow(None, f'{window_text}')
                    if hwnd:
                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)  # maximizes
                        break
                    time.sleep(0.1)
            
            active_window = None
            listener_running = False
            return False  # Stop listener when middle button is released

def on_move(x, y):
    global is_dragging, active_window, cursor_initial_offset
    
    if is_dragging and active_window:
        try:
            # Calculate new window position to keep cursor at center
            new_x = x - active_window.width // 2 - cursor_initial_offset[0]
            new_y = y - active_window.height // 2 - cursor_initial_offset[1]
            
            # Move the window
            active_window.moveTo(int(new_x), int(new_y))
            
        except Exception as e:
            print(f"Error while moving: {e}")

def startListener():
    global listener_running
    
    # Reset variables
    listener_running = True
    
    # Create listeners
    mouse_listener = mouse.Listener(on_click=on_click, on_move=on_move)
    
    # Start listeners
    mouse_listener.start()
    
    print("Listening...")
    
    # Wait until a signal is received
    while listener_running:
        time.sleep(0.01)
    
    # Stop both listeners
    mouse_listener.stop()
    
    # Wait for listeners to actually stop
    mouse_listener.join()

while True:
    startListener()