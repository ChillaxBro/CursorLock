import win32gui
import ctypes
import time
import threading
import keyboard
from win10toast import ToastNotifier
import pystray
from PIL import Image
import os
import sys
import win32con
import hashlib
from packaging import version
import platform

class CursorLocker:
    def __init__(self):
        self.locked = False
        self.lock_thread = None
        self.current_window = "None"
        self.locked_hwnd = None
        self.notifier = None  # Will be initialized when needed
        self.icon = None
        self.last_notification_time = 0
        self.force_unlock_flag = False
        self.current_hotkey = "ctrl+alt+l"
        self.hotkey_recording = False
        self.recorded_keys = set()
        self.keyboard_hook_id = None
        self.active_notification = None
        
        # Ensure single instance
        self.mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "CursorLockerMutex")
        if ctypes.windll.kernel32.GetLastError() == 183:
            sys.exit(0)

    def get_resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        # Handle both .ico and .png files
        for ext in ['.ico', '.png']:
            path = os.path.join(base_path, relative_path + ext)
            if os.path.exists(path):
                return path
        
        return os.path.join(base_path, relative_path)

    def show_notification(self, title, message, duration=2):
        """Show instant notification with controlled duration"""
        try:
            # Ensure we have a notifier instance
            if not hasattr(self, 'notifier') or self.notifier is None:
                self.notifier = ToastNotifier()
            
            # Windows 10+ requires a unique message ID to show multiple notifications
            notification_id = hash(f"{title}{message}") % 10000
            
            # Windows notifications sometimes need a small delay between messages
            time_since_last = time.time() - self.last_notification_time
            if time_since_last < 0.5:
                time.sleep(0.5 - time_since_last)
            
            # Get icon path if available
            icon_path = None
            try:
                icon_path = self.get_resource_path("icon")
                if not os.path.exists(icon_path):
                    icon_path = None
            except:
                pass
            
            # Show notification
            self.notifier.show_toast(
                title=title,
                msg=message,
                duration=duration,
                threaded=True,
                icon_path=icon_path
            )
            self.last_notification_time = time.time()
            print(f"Notification shown: {title} - {message}")
        except Exception as e:
            print(f"Notification failed: {str(e)}")
            # Fallback to simple message box if notifications fail
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x40)

    def start_hotkey_recording(self):
        """Start recording a new hotkey combination"""
        if self.hotkey_recording:
            return
            
        self.hotkey_recording = True
        self.recorded_keys = set()
        self.show_notification("Hotkey Setup", "Recording new hotkey... Press your desired key combination", duration=2)
        
        # Store the hook ID so we can properly remove it later
        self.keyboard_hook_id = keyboard.hook(self.keyboard_hook)

    def keyboard_hook(self, event):
        """Handle key events during hotkey recording"""
        if not self.hotkey_recording:
            return
            
        if event.event_type == keyboard.KEY_DOWN:
            # Standardize key names and handle modifiers
            key_name = event.name.lower()
            if key_name in ['ctrl', 'alt', 'shift', 'windows']:
                key_name = f"{key_name}"  # Remove _left/_right distinction
            
            # Add to recorded keys
            if key_name not in self.recorded_keys:
                self.recorded_keys.add(key_name)
                
        elif event.event_type == keyboard.KEY_UP:
            # If any key is released and we have keys recorded, finish
            if self.recorded_keys:
                self.finish_hotkey_recording()

    def finish_hotkey_recording(self):
        """Finish recording and set new hotkey"""
        if not self.hotkey_recording:
            return
            
        try:
            # Clean up the keyboard hook
            if self.keyboard_hook_id is not None:
                keyboard.unhook(self.keyboard_hook_id)
                self.keyboard_hook_id = None
            
            if self.recorded_keys:
                # Remove old hotkey if it exists
                try:
                    keyboard.remove_hotkey(self.current_hotkey)
                except:
                    pass
                
                # Create new hotkey string (sorted for consistency)
                self.current_hotkey = '+'.join(sorted(self.recorded_keys))
                
                # Add new hotkey
                keyboard.add_hotkey(self.current_hotkey, self.toggle_cursor_lock)
                
                # Update tray menu
                if self.icon:
                    self.update_tray_menu()
                
                self.show_notification("Hotkey Saved", f"New hotkey set: {self.current_hotkey}", duration=2)
            else:
                self.show_notification("Hotkey Setup", "No keys recorded", duration=2)
        except Exception as e:
            self.show_notification("Hotkey Error", f"Failed to set hotkey: {str(e)}", duration=3)
        finally:
            self.hotkey_recording = False
            self.recorded_keys = set()

    def update_tray_menu(self):
        """Update the tray menu with current settings"""
        if not self.icon:
            return
            
        # Try to load icon again in case it wasn't loaded properly initially
        try:
            icon_path = self.get_resource_path("icon")
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                image = Image.new('RGB', (64, 64), (70, 130, 180))  # Fallback
        except:
            image = Image.new('RGB', (64, 64), (70, 130, 180))
        
        menu = pystray.Menu(
            pystray.MenuItem(
                f'Current Hotkey: {self.current_hotkey}',
                lambda: None  # Non-clickable item
            ),
            pystray.MenuItem(
                'Set New Hotkey',
                lambda: self.start_hotkey_recording()
            ),
            pystray.MenuItem(
                'Toggle Lock',
                lambda: self.toggle_cursor_lock()
            ),
            pystray.MenuItem(
                'Force Unlock',
                lambda: self.force_unlock()
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                'Exit',
                lambda: self.exit_program()
            )
        )
        
        self.icon.menu = menu
        self.icon.update_menu()

    def get_window_info(self, hwnd):
        """Get window information for specific handle"""
        try:
            if hwnd and win32gui.IsWindowVisible(hwnd):
                rect = win32gui.GetWindowRect(hwnd)
                title = win32gui.GetWindowText(hwnd)
                return rect, title or "Untitled Window"
            return None, "Invalid window"
        except Exception as e:
            return None, f"Error: {str(e)}"

    def apply_cursor_lock(self, rect):
        """Lock cursor to specified rectangle"""
        try:
            left, top, right, bottom = rect
            if right - left > 10 and bottom - top > 10:
                ctypes.windll.user32.ClipCursor(ctypes.byref(ctypes.wintypes.RECT(left, top, right, bottom)))
        except Exception as e:
            print(f"Lock error: {e}")

    def force_unlock(self):
        """Forcefully unlock cursor and clean up any residual state"""
        try:
            ctypes.windll.user32.ClipCursor(None)
            self.locked = False
            self.current_window = "None"
            self.locked_hwnd = None
            self.force_unlock_flag = False
            
            if self.lock_thread and self.lock_thread.is_alive():
                self.lock_thread.join(0.1)
                
            print("Forcefully unlocked cursor and reset state")
        except Exception as e:
            print(f"Force unlock error: {e}")

    def lock_loop(self):
        """Main locking loop that maintains lock on original window"""
        try:
            original_rect, original_title = self.get_window_info(self.locked_hwnd)
            if not original_rect:
                self.show_notification("Lock Error", "Original window not available", duration=3)
                self.force_unlock_flag = True
                return

            while self.locked and not self.force_unlock_flag:
                self.apply_cursor_lock(original_rect)
                current_hwnd = win32gui.GetForegroundWindow()
                if current_hwnd != self.locked_hwnd:
                    try:
                        win32gui.SetForegroundWindow(self.locked_hwnd)
                    except:
                        pass
                time.sleep(0.05)
        finally:
            self.force_unlock_flag = True
            self.force_unlock()

    def toggle_cursor_lock(self):
        """Toggle lock state with complete reset when deactivating"""
        if not self.locked:
            hwnd, rect, title = self.get_active_window_info()
            if not hwnd:
                return
                
            self.locked = True
            self.force_unlock_flag = False
            self.locked_hwnd = hwnd
            self.current_window = title[:60] + (title[60:] and '...')
            self.lock_thread = threading.Thread(target=self.lock_loop, daemon=True)
            self.lock_thread.start()
            self.show_notification("Cursor Lock", "ACTIVATED - Cursor locked to current window", duration=2)
        else:
            self.force_unlock_flag = True
            self.show_notification("Cursor Lock", "DEACTIVATED - Cursor unlocked", duration=2)

    def get_active_window_info(self):
        """Get active window information"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd and win32gui.IsWindowVisible(hwnd):
                rect = win32gui.GetWindowRect(hwnd)
                title = win32gui.GetWindowText(hwnd)
                return hwnd, rect, title or "Untitled Window"
            return None, None, "No active window"
        except Exception as e:
            return None, None, f"Error: {str(e)}"

    def create_tray_icon(self):
        """Create system tray icon with menu"""
        try:
            # Try to load icon from resources
            icon_path = self.get_resource_path("icon")
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                # Fallback to default icon
                image = Image.new('RGB', (64, 64), (70, 130, 180))
            
            menu = pystray.Menu(
                pystray.MenuItem(
                    f'Current Hotkey: {self.current_hotkey}',
                    lambda: None
                ),
                pystray.MenuItem(
                    'Set New Hotkey',
                    lambda: self.start_hotkey_recording()
                ),
                pystray.MenuItem(
                    'Toggle Lock',
                    lambda: self.toggle_cursor_lock()
                ),
                pystray.MenuItem(
                    'Force Unlock',
                    lambda: self.force_unlock()
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    'Exit',
                    lambda: self.exit_program()
                )
            )
            
            self.icon = pystray.Icon(
                "cursor_lock",
                image,
                "Cursor Lock to Window",
                menu
            )
            return self.icon
        except Exception as e:
            print(f"Tray icon error: {e}")
            return None

    def exit_program(self):
        """Clean exit with full state reset"""
        self.force_unlock_flag = True
        self.force_unlock()
        if self.icon:
            try:
                self.icon.stop()
            except:
                pass
        if self.keyboard_hook_id is not None:
            keyboard.unhook(self.keyboard_hook_id)
        os._exit(0)

    def run(self):
        """Main application entry"""
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
        self.show_notification("Cursor Lock", f"Service started\nPress {self.current_hotkey} to toggle lock", duration=3)
        keyboard.add_hotkey(self.current_hotkey, self.toggle_cursor_lock)
        
        icon = self.create_tray_icon()
        try:
            if icon:
                icon.run()
            else:
                while True:
                    time.sleep(1)
        except KeyboardInterrupt:
            self.exit_program()
        except Exception as e:
            print(f"Runtime error: {e}")
            self.exit_program()

if __name__ == "__main__":
    app = CursorLocker()
    app.run()

#"pyinstaller --onefile --windowed --icon=icon.ico --add-data "icon.ico;." --hidden-import=pystray._win32 --hidden-import=win10toast CursorLock.py"