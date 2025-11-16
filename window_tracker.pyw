# ============================================
# WINDOW TIME TRACKER WITH PROGRAM FILTERING
# ============================================

import win32gui
import win32process
import psutil
import time
import csv
import os
import re
from datetime import datetime
from collections import defaultdict
import threading
import pystray
from PIL import Image, ImageDraw
import atexit
from ctypes import Structure, windll, c_uint, sizeof, byref
import subprocess

# ============================================
# IDLE TIME DETECTION (Windows)
# ============================================

class LASTINPUTINFO(Structure):
    """
    Structure to hold Windows last input info
    Used to detect when user is idle (no mouse/keyboard activity)
    """
    _fields_ = [
        ('cbSize', c_uint),
        ('dwTime', c_uint),
    ]

def get_idle_duration():
    """
    Returns how long (in seconds) the system has been idle
    Idle = no mouse movement or keyboard input
    """
    lastInputInfo = LASTINPUTINFO()
    lastInputInfo.cbSize = sizeof(lastInputInfo)
    if windll.user32.GetLastInputInfo(byref(lastInputInfo)):
        millis = windll.kernel32.GetTickCount() - lastInputInfo.dwTime
        return millis / 1000.0
    return 0

# ============================================
# DEFAULT CONFIGURATION (used if file missing)
# ============================================

DEFAULT_CHECK_INTERVAL = 20
DEFAULT_LOG_FOLDER = "window_logs"
DEFAULT_IDLE_THRESHOLD = 300  # 5 minutes in seconds
DEFAULT_AUTO_SAVE_INTERVAL = 1800  # NEW: 30 minutes in seconds
CONFIG_FILE = "tracker_config.txt"

# ============================================
# CONFIGURATION LOADER
# ============================================

def load_configuration():
    """
    Loads all settings from the configuration file
    Returns: (check_interval, log_folder, idle_threshold, auto_save_interval, tracked_programs, ignored_programs)
    """
    # Default values
    check_interval = DEFAULT_CHECK_INTERVAL
    log_folder = DEFAULT_LOG_FOLDER
    idle_threshold = DEFAULT_IDLE_THRESHOLD
    auto_save_interval = DEFAULT_AUTO_SAVE_INTERVAL  # NEW: Auto-save default
    tracked_programs = set()
    ignored_programs = set()
    
    # Create config file if it doesn't exist
    if not os.path.exists(CONFIG_FILE):
        create_default_config()
    
    # Read the configuration file
    current_section = None
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                
                # Skip empty lines and comments
                if not line or line.startswith('#'):
                    continue
                
                # Check for section headers
                if line.startswith('[') and line.endswith(']'):
                    current_section = line[1:-1].upper()
                    continue
                
                # Parse settings (key = value)
                if '=' in line and current_section is None:
                    key, value = line.split('=', 1)
                    key = key.strip().lower()
                    value = value.strip()
                    
                    if key == 'check_interval':
                        try:
                            check_interval = int(value)
                            # Allow range from 5 seconds to 1 hour (3600 seconds)
                            if check_interval < 5:
                                check_interval = 5
                                print(f"Warning: check_interval too low, set to 5 seconds")
                            elif check_interval > 3600:
                                check_interval = 3600
                                print(f"Warning: check_interval too high, set to 1 hour")
                        except ValueError:
                            check_interval = DEFAULT_CHECK_INTERVAL
                            print(f"Warning: Invalid check_interval value, using default: {DEFAULT_CHECK_INTERVAL}")
                    
                    elif key == 'log_folder':
                        log_folder = value
                    
                    elif key == 'idle_threshold':
                        try:
                            idle_threshold = int(value)
                            # Allow range from 5 second to 1 hour (3600 seconds)
                            if idle_threshold < 5:
                                idle_threshold = 5
                                print(f"Warning: idle_threshold too low, set to 5 seconds")
                            elif idle_threshold > 3600:
                                idle_threshold = 3600
                                print(f"Warning: idle_threshold too high, set to 1 hour")
                        except ValueError:
                            idle_threshold = DEFAULT_IDLE_THRESHOLD
                            print(f"Warning: Invalid idle_threshold value, using default: {DEFAULT_IDLE_THRESHOLD}")
                    
                    # NEW: Auto-save interval configuration
                    elif key == 'auto_save_interval':
                        try:
                            auto_save_interval = int(value)
                            # Allow range from 1 second to 24 hours (86400 seconds), or 0 to disable
                            if auto_save_interval < 0:
                                auto_save_interval = 0
                                print(f"Warning: auto_save_interval negative, disabling auto-save")
                            elif auto_save_interval > 86400:
                                auto_save_interval = 86400
                                print(f"Warning: auto_save_interval too high, set to 24 hours")
                        except ValueError:
                            auto_save_interval = DEFAULT_AUTO_SAVE_INTERVAL
                            print(f"Warning: Invalid auto_save_interval value, using default: {DEFAULT_AUTO_SAVE_INTERVAL}")
                
                # Parse program lists
                elif current_section == 'TRACKED_PROGRAMS':
                    tracked_programs.add(line.lower())
                elif current_section == 'IGNORED_PROGRAMS':
                    ignored_programs.add(line.lower())
    
    except Exception as e:
        print(f"Error reading config file: {e}")
        print("Using default settings")
    
    # Note: No hardcoded programs anymore - everything is in the config file!
    return check_interval, log_folder, idle_threshold, auto_save_interval, tracked_programs, ignored_programs

def create_default_config():
    """
    Creates a default configuration file with explanations
    """
    default_content = """# ================================================
# WINDOW TIME TRACKER - CONFIGURATION FILE
# ================================================
# Edit these settings to customize your tracker
# Save this file and restart the program for changes to take effect

# ------------------------------------------------
# TIMING SETTINGS
# ------------------------------------------------

# How often to check for window changes (in seconds)
# Lower = more accurate, but uses more computer resources
# Recommended: 15-30 seconds
# Valid range: 5 to 3600 (5 seconds to 1 hour)
check_interval = 20

# Idle threshold (in seconds)
# How long without mouse/keyboard activity before counting as "idle"
# Recommended: 300 seconds (5 minutes) for architecture work
# This allows time for thinking, reviewing prints, or sketching
# Valid range: 5 to 3600 (5 seconds to 1 hour)
idle_threshold = 300

# Auto-save interval (in seconds)
# How often to automatically save active sessions
# Set to 0 to disable auto-save
# Recommended: 1800 seconds (30 minutes)
# Valid range: 1 to 86400 (1 second to 24 hours), or 0 to disable
auto_save_interval = 1800

# Folder where log files are saved
# You can use a full path like C:\\Users\\YourName\\Documents\\ProjectLogs
log_folder = window_logs

# ------------------------------------------------
# PROGRAMS TO TRACK
# ------------------------------------------------
# List one program per line (use the .exe filename)
# To find a program's .exe name:
#   1. Open the program
#   2. Press Ctrl+Shift+Esc (opens Task Manager)
#   3. Go to "Details" tab
#   4. Find your program and look at the "Name" column
# Examples:

[TRACKED_PROGRAMS]
blender.exe
archicad.exe
revit.exe
autocad.exe
sketchup.exe
rhino.exe
3dsmax.exe
lumion.exe
enscape.exe
photoshop.exe
illustrator.exe
vectorworks.exe
# Add your own programs below:


# ------------------------------------------------
# PROGRAMS TO IGNORE
# ------------------------------------------------
# These programs will NEVER be tracked
# Common browsers are listed below by default
# You can remove browsers from this list if you want to track them
# Or add any other programs you want to ignore

[IGNORED_PROGRAMS]
# Web Browsers (commonly ignored)
chrome.exe
firefox.exe
msedge.exe
opera.exe
brave.exe
iexplore.exe
safari.exe

# Windows system programs
notepad.exe
explorer.exe

# Add more programs to ignore below:

"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        f.write(default_content)

# ============================================
# HELPER FUNCTIONS
# ============================================

def get_process_name_from_hwnd(hwnd):
    """
    Gets the actual program name (like blender.exe) from a window handle
    """
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return process.name().lower()
    except:
        return None

def get_all_open_windows():
    """
    Returns list of tuples: (window_title, process_name)
    Only includes windows that appear on the taskbar
    """
    windows = []
    
    def callback(hwnd, extra):
        # Check if window is visible
        if not win32gui.IsWindowVisible(hwnd):
            return
        
        # Get window title
        window_title = win32gui.GetWindowText(hwnd)
        if not window_title:
            return
        
        # Filter out windows that don't appear on taskbar
        # Check if window has WS_EX_APPWINDOW or doesn't have WS_EX_TOOLWINDOW
        ex_style = win32gui.GetWindowLong(hwnd, -20)  # GWL_EXSTYLE = -20
        
        # WS_EX_TOOLWINDOW = 0x00000080 (tool windows don't show on taskbar)
        # WS_EX_APPWINDOW = 0x00040000 (force window on taskbar)
        WS_EX_TOOLWINDOW = 0x00000080
        WS_EX_APPWINDOW = 0x00040000
        
        # Skip tool windows (unless they have APPWINDOW flag)
        if (ex_style & WS_EX_TOOLWINDOW) and not (ex_style & WS_EX_APPWINDOW):
            return
        
        # Check if window has an owner (owned windows usually don't show on taskbar)
        owner = win32gui.GetWindow(hwnd, 4)  # GW_OWNER = 4
        if owner != 0:
            return
        
        # If we got here, this is a taskbar window
        process_name = get_process_name_from_hwnd(hwnd)
        windows.append((window_title, process_name))
    
    win32gui.EnumWindows(callback, None)
    return windows

# FOCUS TIME: New helper function to get the focused window
def get_focused_window_info():
    """
    Returns the window title and process name of the currently focused (active) window
    This is the window that's currently selected in your taskbar
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            window_title = win32gui.GetWindowText(hwnd)
            process_name = get_process_name_from_hwnd(hwnd)
            return window_title, process_name
    except:
        pass
    return None, None

def extract_project_name(window_title):
    """
    Extracts the project name from a window title
    Removes everything between [ and ] brackets (like file paths)
    """
    # Remove everything between [ and ] including the brackets
    cleaned_title = re.sub(r'\[.*?\]', '', window_title).strip()
    
    # Extract project name from cleaned title
    match = re.search(r'^(.+?)\s*[-–—]', cleaned_title)
    if match:
        project_name = match.group(1).strip()
    else:
        project_name = cleaned_title[:50]
    
    # Remove leading asterisks and spaces
    project_name = project_name.lstrip('* ')
    
    # Replace illegal filename characters
    project_name = re.sub(r'[<>:"/\\|?*]', '_', project_name)
    
    if len(project_name) < 2:
        project_name = "Unnamed_Project"
    
    return project_name

def get_app_name_from_title(window_title):
    """
    Tries to identify which application a window belongs to
    """
    title_lower = window_title.lower()
    
    app_keywords = {
        'blender': 'Blender',
        'archicad': 'ArchiCAD',
        'revit': 'Revit',
        'autocad': 'AutoCAD',
        'sketchup': 'SketchUp',
        'rhino': 'Rhino',
        '3ds max': '3dsMax',
        'lumion': 'Lumion',
        'enscape': 'Enscape',
        'photoshop': 'Photoshop',
        'illustrator': 'Illustrator',
        'vectorworks': 'Vectorworks',
    }
    
    for keyword, app_name in app_keywords.items():
        if keyword in title_lower:
            return app_name
    
    parts = window_title.split(' - ')
    if len(parts) > 1:
        return parts[-1][:30]
    
    return "Unknown_App"

# ============================================
# MAIN TRACKING CLASS
# ============================================

class WindowTimeTracker:
    def __init__(self, check_interval, log_folder, idle_threshold, auto_save_interval, tracked_programs, ignored_programs):
        """
        Initializes the window time tracker
        """
        self.check_interval = check_interval
        self.log_folder = log_folder
        self.idle_threshold = idle_threshold
        self.auto_save_interval = auto_save_interval  # NEW: Store auto-save interval
        self.tracked_programs = tracked_programs
        self.ignored_programs = ignored_programs
        
        self.session_active = {}
        self.sessions = defaultdict(list)
        self.log_meta = {}
        
        # NEW: Track which sessions have been saved manually and their index
        self.saved_session_index = {}  # Maps project -> index in sessions list
        
        self.running = True
        self.stop_event = threading.Event()
        
        # NEW: Auto-save timer thread
        self.auto_save_thread = None
        self.auto_save_stop_event = threading.Event()
        
        # Create log folder if it doesn't exist
        if not os.path.exists(self.log_folder):
            os.makedirs(self.log_folder)
        
        self.load_existing_sessions()
        
        # NEW: Start auto-save thread if enabled
        if self.auto_save_interval > 0:
            self.start_auto_save()

    # NEW: Auto-save thread management
    def start_auto_save(self):
        """
        Starts the auto-save background thread
        """
        print(f"Auto-save enabled: will save every {self.auto_save_interval} seconds ({self.auto_save_interval/60:.1f} minutes)")
        self.auto_save_thread = threading.Thread(target=self._auto_save_loop, daemon=False)
        self.auto_save_thread.start()
    
    def _auto_save_loop(self):
        """
        Background loop that triggers auto-save at regular intervals
        """
        while not self.auto_save_stop_event.is_set():
            # Wait for the auto-save interval (or until stop signal)
            self.auto_save_stop_event.wait(self.auto_save_interval)
            
            # If we were interrupted by stop signal, exit
            if self.auto_save_stop_event.is_set():
                break
            
            # Trigger auto-save
            print(f"\n[AUTO-SAVE] Running automatic save at {datetime.now().strftime('%H:%M:%S')}")
            try:
                saved_count = self.manual_save_all_logs()
                print(f"[AUTO-SAVE] Completed: saved {saved_count} session(s)")
                
                # NEW: Notify user via system tray (if icon is available)
                global icon
                if icon and saved_count > 0:
                    icon.notify(
                        title="Auto-Save Complete",
                        message=f"Saved {saved_count} active session(s) at {datetime.now().strftime('%H:%M')}"
                    )
            except Exception as e:
                print(f"[AUTO-SAVE] Error: {e}")

    def should_track_window(self, process_name):
        """
        Decides if a window should be tracked based on its program name
        """
        if not process_name:
            return False
        
        process_name = process_name.lower()
        
        # Never track ignored programs
        if process_name in self.ignored_programs:
            return False
        
        # Only track programs in our tracked list
        return process_name in self.tracked_programs

    def load_existing_sessions(self):
        """
        Loads previously saved session data from log files
        """
        if not os.path.exists(self.log_folder):
            return
        
        for fname in os.listdir(self.log_folder):
            if fname.endswith('_log.csv'):
                proj = fname[:-8]
                filepath = os.path.join(self.log_folder, fname)
                
                with open(filepath, 'r', encoding='utf-8') as f:
                    meta = {}
                    for line in f:
                        if line.startswith('# Created:'):
                            meta['created'] = line.strip().split(':', 1)[1].strip()
                        elif line.startswith('# Last updated:'):
                            meta['last_changed'] = line.strip().split(':', 1)[1].strip()
                        elif not line.startswith('#') and line.strip():
                            break
                    
                    self.log_meta[proj] = meta if meta else {
                        'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'last_changed': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                
                with open(filepath, 'r', encoding='utf-8') as f:
                    reader = csv.reader(filter(lambda row: row[0] != '#', f))
                    next(reader, None)
                    for row in reader:
                        if len(row) >= 3:
                            # FOCUS TIME: Load focus_time if it exists (backward compatible)
                            self.sessions[proj].append({
                                'start': row[0],
                                'end': row[1],
                                'duration': int(row[2]),
                                'idle_time': int(row[3]) if len(row) > 3 else 0,
                                'focus_time': int(row[4]) if len(row) > 4 else 0,  # NEW: focus_time column
                                'active_time': int(row[5]) if len(row) > 5 else 0,
                                'window_title': row[6] if len(row) > 6 else ''
                            })

    def on_window_open(self, window_title):
        """
        Called when a new tracked window is detected
        """
        proj = extract_project_name(window_title)
        if proj not in self.session_active:
            self.session_active[proj] = {
                'start': datetime.now(),
                'window_title': window_title,
                'idle_time': 0,
                'focus_time': 0,  # FOCUS TIME: Initialize focus time tracking
                'was_idle': False,
                'last_check_time': time.time()  # FOCUS TIME: Track time between checks
            }

    def on_window_close(self, window_title):
        """
        Called when a tracked window closes
        """
        proj = extract_project_name(window_title)
        if proj in self.session_active:
            session = self.session_active.pop(proj)
            end_time = datetime.now()
            duration = int((end_time - session['start']).total_seconds())
            idle_time = session['idle_time']
            focus_time = session['focus_time']  # FOCUS TIME: Get tracked focus time
            active_time = duration - idle_time
            
            # NEW: Check if this session was already saved manually
            if proj in self.saved_session_index:
                # UPDATE the existing session instead of creating a new one
                session_idx = self.saved_session_index[proj]
                self.sessions[proj][session_idx]['end'] = end_time.strftime('%Y-%m-%d %H:%M:%S')
                self.sessions[proj][session_idx]['duration'] = duration
                self.sessions[proj][session_idx]['idle_time'] = idle_time
                self.sessions[proj][session_idx]['focus_time'] = focus_time  # FOCUS TIME: Save it
                self.sessions[proj][session_idx]['active_time'] = active_time
                
                # Clear the saved session tracking
                del self.saved_session_index[proj]
                print(f"Updated saved session for {proj}")
            else:
                # No manual save was done, create a new session normally
                self.sessions[proj].append({
                    'start': session['start'].strftime('%Y-%m-%d %H:%M:%S'),
                    'end': end_time.strftime('%Y-%m-%d %H:%M:%S'),
                    'duration': duration,
                    'idle_time': idle_time,
                    'focus_time': focus_time,  # FOCUS TIME: Save it
                    'active_time': active_time,
                    'window_title': window_title
                })
            
            self.save_session_log(proj)

    def save_session_log(self, proj):
        """
        Saves session data to a CSV file with improved formatting
        """
        if not self.sessions[proj]:
            return
        
        fname = os.path.join(self.log_folder, f"{proj}_log.csv")
        
        # Calculate totals
        total_time = sum(s['duration'] for s in self.sessions[proj])
        total_idle = sum(s.get('idle_time', 0) for s in self.sessions[proj])
        total_focus = sum(s.get('focus_time', 0) for s in self.sessions[proj])  # FOCUS TIME: Calculate total
        total_active = total_time - total_idle
        
        created = self.log_meta.get(proj, {}).get('created', self.sessions[proj][0]['start'])
        last_changed = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_meta[proj] = {'created': created, 'last_changed': last_changed}
        
        # Calculate work efficiency
        work_efficiency = (total_active / total_time * 100) if total_time > 0 else 0
        focus_percentage = (total_focus / total_time * 100) if total_time > 0 else 0  # FOCUS TIME: Calculate percentage
        
        # Create nicely formatted header
        header_lines = [
            f'# Project: {proj}',
            f'# Created: {created}',
            f'# Last updated: {last_changed}',
            f'# Number of sessions: {len(self.sessions[proj])}',
            f'#',
            f'# ===== TIME SUMMARY =====',
            f'# Total time (all sessions): {total_time} sec ({total_time/3600:.2f} hours)',
            f'# Active work time: {total_active} sec ({total_active/3600:.2f} hours)',
            f'# Focus time (window active): {total_focus} sec ({total_focus/3600:.2f} hours)',  # FOCUS TIME: Show in summary
            f'# Idle time: {total_idle} sec ({total_idle/3600:.2f} hours)',
            f'# Work efficiency: {work_efficiency:.1f}%',
            f'# Focus percentage: {focus_percentage:.1f}%',  # FOCUS TIME: Show percentage
            f'#',
            f'# ===== SESSION DATA ====='
        ]
        
        with open(fname, 'w', encoding='utf-8', newline='') as f:
            # Write header
            for line in header_lines:
                f.write(line + '\n')
            
            writer = csv.writer(f)
            
            # Write column headers - FOCUS TIME: Added focus_time_sec column before active_time_sec
            writer.writerow([
                'session_start',
                'session_end',
                'session_duration_sec',
                'idle_time_sec',
                'focus_time_sec',  # FOCUS TIME: New column (placed before active_time_sec as requested)
                'active_time_sec',
                'window_title'
            ])
            
            # Write data rows
            for s in self.sessions[proj]:
                writer.writerow([
                    s['start'],
                    s['end'],
                    s['duration'],
                    s.get('idle_time', 0),
                    s.get('focus_time', 0),  # FOCUS TIME: Write focus time data
                    s.get('active_time', s['duration'] - s.get('idle_time', 0)),
                    s['window_title']
                ])

    def save_all_active_sessions(self):
        """
        Saves all currently running sessions
        This is called when the program exits to prevent data loss
        """
        print("Saving all active sessions...")
        for proj in list(self.session_active.keys()):
            wtitle = self.session_active[proj]['window_title']
            self.on_window_close(wtitle)
        print("All active sessions saved!")

    def manual_save_all_logs(self):
        """
        FIXED: Manually saves all active sessions as current snapshots
        The session will be UPDATED (not duplicated) when the window closes
        """
        print("Manual save requested - saving current state...")
        saved_count = 0
        
        for proj in list(self.session_active.keys()):
            session = self.session_active[proj]
            end_time = datetime.now()
            duration = int((end_time - session['start']).total_seconds())
            idle_time = session['idle_time']
            focus_time = session['focus_time']  # FOCUS TIME: Get current focus time
            active_time = duration - idle_time
            
            # Check if we already saved this session manually before
            if proj in self.saved_session_index:
                # UPDATE the existing saved session
                session_idx = self.saved_session_index[proj]
                self.sessions[proj][session_idx]['end'] = end_time.strftime('%Y-%m-%d %H:%M:%S')
                self.sessions[proj][session_idx]['duration'] = duration
                self.sessions[proj][session_idx]['idle_time'] = idle_time
                self.sessions[proj][session_idx]['focus_time'] = focus_time  # FOCUS TIME: Update it
                self.sessions[proj][session_idx]['active_time'] = active_time
                print(f"Updated existing saved session for {proj}")
            else:
                # First manual save for this session - create NEW entry
                self.sessions[proj].append({
                    'start': session['start'].strftime('%Y-%m-%d %H:%M:%S'),
                    'end': end_time.strftime('%Y-%m-%d %H:%M:%S'),
                    'duration': duration,
                    'idle_time': idle_time,
                    'focus_time': focus_time,  # FOCUS TIME: Save it
                    'active_time': active_time,
                    'window_title': session['window_title']
                })
                
                # Remember this session index so we can update it later
                self.saved_session_index[proj] = len(self.sessions[proj]) - 1
                print(f"Created new saved session for {proj} at index {self.saved_session_index[proj]}")
            
            # Write to file
            self.save_session_log(proj)
            saved_count += 1
        
        # DO NOT reset session start time - we keep tracking from the original start
        # This ensures when window closes, we update with the full duration
        
        print(f"Manual save complete! Saved {saved_count} active session(s).")
        return saved_count

    def update_idle_time(self):
        """
        Updates idle time for all active sessions
        Called periodically to track idle periods
        """
        current_idle = get_idle_duration()
        
        for proj in self.session_active:
            session = self.session_active[proj]
            
            # If idle time exceeds threshold, we're in idle state
            if current_idle >= self.idle_threshold:
                if not session['was_idle']:
                    # Just became idle, mark it
                    session['was_idle'] = True
                    session['idle_start'] = current_idle
                else:
                    # Continue being idle, accumulate time
                    idle_to_add = current_idle - session.get('idle_start', current_idle)
                    session['idle_time'] += int(idle_to_add)
                    session['idle_start'] = current_idle
            else:
                # Not idle anymore (or never was)
                session['was_idle'] = False
                if 'idle_start' in session:
                    del session['idle_start']

    # FOCUS TIME: New method to update focus time for active sessions
    def update_focus_time(self):
        """
        Updates focus time for the currently focused window
        This tracks when a tracked window is the active window (selected in taskbar)
        """
        focused_title, focused_process = get_focused_window_info()
        current_time = time.time()
        
        # Check each active session
        for proj in self.session_active:
            session = self.session_active[proj]
            time_since_last_check = current_time - session.get('last_check_time', current_time)
            
            # Check if this session's window is the currently focused window
            if focused_title and extract_project_name(focused_title) == proj:
                # This window is focused! Add the time since last check
                session['focus_time'] += int(time_since_last_check)
            
            # Update last check time for next iteration
            session['last_check_time'] = current_time

    def run(self):
        """
        Main tracking loop
        """
        try:
            while not self.stop_event.is_set():
                current_windows = get_all_open_windows()
                tracked_now = set()
                
                for wtitle, process_name in current_windows:
                    if self.should_track_window(process_name):
                        proj = extract_project_name(wtitle)
                        tracked_now.add(proj)
                        if proj not in self.session_active:
                            self.on_window_open(wtitle)
                
                # Close sessions for windows that are no longer open
                closed_projs = set(self.session_active) - tracked_now
                for proj in list(closed_projs):
                    wtitle = self.session_active[proj]['window_title']
                    self.on_window_close(wtitle)
                
                # Update idle time for all active sessions
                self.update_idle_time()
                
                # FOCUS TIME: Update focus time for all active sessions
                self.update_focus_time()
                
                # Use wait() instead of sleep() so we can interrupt immediately
                self.stop_event.wait(self.check_interval)
        
        finally:
            print("Tracker shutting down...")
            self.save_all_active_sessions()

    def stop(self):
        """
        Stops the tracking loop gracefully
        """
        print("Stop requested...")
        self.running = False
        self.stop_event.set()
        
        # NEW: Stop auto-save thread
        if self.auto_save_thread:
            print("Stopping auto-save thread...")
            self.auto_save_stop_event.set()
            self.auto_save_thread.join(timeout=5)

# ============================================
# SYSTEM TRAY ICON
# ============================================

def create_tray_icon():
    """
    Creates a simple icon for the system tray
    """
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), (255, 255, 255))
    dc = ImageDraw.Draw(image)
    
    dc.ellipse([8, 8, 56, 56], fill=(0, 120, 215), outline=(0, 90, 180))
    dc.rectangle([30, 16, 34, 32], fill=(255, 255, 255))
    dc.rectangle([30, 32, 46, 36], fill=(255, 255, 255))
    
    return image

def on_view_reports(icon, item):
    """
    Opens the log folder in Windows Explorer
    """
    global tracker
    log_path = os.path.abspath(tracker.log_folder)
    
    # Create folder if it doesn't exist
    if not os.path.exists(log_path):
        os.makedirs(log_path)
    
    # Open in Windows Explorer
    subprocess.Popen(f'explorer "{log_path}"')

def on_save_all_logs(icon, item):
    """
    Manually saves all currently active logs
    Called when user clicks "Save All Logs Now" in the system tray
    """
    global tracker
    try:
        saved_count = tracker.manual_save_all_logs()
        
        # Show notification to user
        if saved_count > 0:
            icon.notify(
                title="Logs Saved!",
                message=f"Saved {saved_count} active session(s). Tracking continues."
            )
        else:
            icon.notify(
                title="Nothing to Save",
                message="No active sessions found."
            )
    except Exception as e:
        print(f"Error during manual save: {e}")
        icon.notify(
            title="Save Error",
            message="An error occurred while saving logs."
        )

def on_quit(icon, item):
    """
    Called when user clicks "Exit" in the system tray
    """
    global tracker, tracking_thread
    print("Exit button clicked...")
    
    # Tell the tracker to stop
    tracker.stop()
    
    # Wait for the tracking thread to finish (with timeout)
    print("Waiting for tracker thread to finish...")
    tracking_thread.join(timeout=5)
    
    print("Stopping system tray icon...")
    # Now stop the icon
    icon.stop()

def emergency_cleanup():
    """
    Backup cleanup in case something goes wrong
    """
    global tracker
    if tracker:
        print("Emergency cleanup triggered...")
        tracker.save_all_active_sessions()

# ============================================
# PROGRAM START
# ============================================

if __name__ == "__main__":
    # Load configuration from file
    check_interval, log_folder, idle_threshold, auto_save_interval, tracked_programs, ignored_programs = load_configuration()
    
    # Create tracker with loaded configuration (including auto_save_interval)
    tracker = WindowTimeTracker(check_interval, log_folder, idle_threshold, auto_save_interval, tracked_programs, ignored_programs)
    
    # Register emergency cleanup
    atexit.register(emergency_cleanup)
    
    # Start tracking in a NON-DAEMON thread
    tracking_thread = threading.Thread(target=tracker.run, daemon=False)
    tracking_thread.start()
    
    # Create system tray icon with menu
    icon_image = create_tray_icon()
    icon = pystray.Icon(
        "Window Tracker",
        icon_image,
        "Window Time Tracker - Running",
        menu=pystray.Menu(
            pystray.MenuItem("View Reports", on_view_reports),
            pystray.MenuItem("Save All Logs Now", on_save_all_logs),
            pystray.MenuItem("Exit", on_quit)
        )
    )
    
    # Run the tray icon
    icon.run()
    
    # Wait for tracking thread to finish
    print("Icon closed, waiting for final cleanup...")
    tracking_thread.join()
    
    print("Program exiting cleanly!")
