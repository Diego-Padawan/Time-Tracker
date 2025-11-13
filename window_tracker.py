# ============================================
# WINDOW TIME TRACKER FOR WINDOWS 10/11
# ============================================
# This program tracks how long different windows/programs are open on your computer
# It works with applications like Blender, ArchiCAD, Revit, etc.

import win32gui  # Library to interact with Windows windows
import win32process  # Library to get process information
import psutil  # Library to get application names
import time  # Library to track time
import csv  # Library to save data in CSV format
import os  # Library to work with files and folders
import re  # Library to work with text patterns
from datetime import datetime  # Library to get current date/time
from collections import defaultdict  # Library for easy dictionary handling

# ============================================
# CONFIGURATION SECTION
# ============================================
# You can modify these settings to customize the program

CHECK_INTERVAL = 20  # How often to check windows (in seconds). Lower = more accurate, but uses more resources
LOG_FOLDER = "window_logs"  # Folder name where logs will be saved
SAVE_INTERVAL = 60  # How often to save data to file (in seconds)

# ============================================
# FUNCTION 1: GET ALL OPEN WINDOWS
# ============================================
# This function finds ALL windows (both visible and background)
# and returns their names/titles

def get_all_open_windows():
    """
    Gets a list of all open window titles on your computer.
    This includes both active (foreground) and inactive (background) windows.
    
    Returns:
        A list of window titles (strings)
    """
    windows = []  # Empty list to store window information
    
    # This inner function is called for each window found
    def callback(hwnd, extra):
        # hwnd = window handle (a unique ID for each window)
        # Check if the window is visible (has a visible window, even if minimized)
        if win32gui.IsWindowVisible(hwnd):
            # Get the window's title text
            window_title = win32gui.GetWindowText(hwnd)
            # Only add windows that have a title (ignore empty ones)
            if window_title:
                windows.append(window_title)
    
    # EnumWindows goes through ALL windows and calls our callback function for each
    win32gui.EnumWindows(callback, None)
    return windows

# ============================================
# FUNCTION 2: EXTRACT PROJECT NAME
# ============================================
# This function tries to find the project name from a window title
# For example: "Project X - Blender 4.5" → extracts "Project X"

def extract_project_name(window_title):
    """
    Extracts the project name from a window title.
    
    Examples:
        "My House - Blender 4.5" → "My House"
        "Building123 - ArchiCAD 26" → "Building123"
        "Untitled - Notepad" → "Untitled"
        "* PROJECTX - Blender" → "PROJECTX"  (strips leading asterisks and spaces)
    
    Args:
        window_title: The full window title string
        
    Returns:
        The extracted project name (string)
    """
    # Common patterns: project name usually comes before " - " separator
    # This pattern looks for text before " - "
    match = re.search(r'^(.+?)\s*[-–—]\s*', window_title)
    
    if match:
        project_name = match.group(1).strip()
    else:
        # If no separator found, use the first part (up to 50 characters)
        project_name = window_title[:50]
    
    # CHANGE 1: Strip leading asterisks and spaces to prevent duplicate log files
    # This makes "* PROJECTX" and "PROJECTX" use the same log file
    project_name = project_name.lstrip('* ')
    
    # Remove special characters that can't be used in filenames
    project_name = re.sub(r'[<>:"/\\|?*]', '_', project_name)
    
    # If the project name is empty or very short, use a default
    if len(project_name) < 2:
        project_name = "Unnamed_Project"
    
    return project_name

# ============================================
# FUNCTION 3: GET APPLICATION NAME
# ============================================
# This function gets the application name (e.g., "Blender", "ArchiCAD")
# from a window title

def get_app_name_from_title(window_title):
    """
    Tries to extract the application name from the window title.
    
    Args:
        window_title: The full window title string
        
    Returns:
        The application name (string)
    """
    # Convert to lowercase for easier matching
    title_lower = window_title.lower()
    
    # Common application keywords to look for
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
        'chrome': 'Chrome',
        'firefox': 'Firefox',
        'excel': 'Excel',
        'word': 'Word',
        'powerpoint': 'PowerPoint',
    }
    
    # Check if any keyword is in the title
    for keyword, app_name in app_keywords.items():
        if keyword in title_lower:
            return app_name
    
    # If no keyword found, try to get it from the end of the title
    # (many apps put their name at the end)
    parts = window_title.split(' - ')
    if len(parts) > 1:
        return parts[-1][:30]  # Return last part (max 30 chars)
    
    return "Unknown_App"

# ============================================
# MAIN TRACKING CLASS
# ============================================
# This class manages all the tracking logic

import win32gui
import win32process
import psutil
import time
import csv
import os
import re
from datetime import datetime
from collections import defaultdict

# ...[ imports and helper functions remain unchanged ]...

class WindowTimeTracker:
    def __init__(self):
        # Store active sessions: {project_name: {'start': datetime, 'window_title': str}}
        self.session_active = {}
        # List of session logs per project: {project_name: [session_dict, ...]}
        self.sessions = defaultdict(list)
        # Log file metadata: {project_name: {'created': str, 'last_changed': str}}
        self.log_meta = {}
        # Ensure log folder exists
        if not os.path.exists(LOG_FOLDER):
            os.makedirs(LOG_FOLDER)
        self.load_existing_sessions()
        print("Window Time Tracker started! Logging by session.")

    def load_existing_sessions(self):
        """Load sessions and log metadata."""
        if not os.path.exists(LOG_FOLDER):
            return
        for fname in os.listdir(LOG_FOLDER):
            if fname.endswith('_log.csv'):
                proj = fname[:-8]
                with open(os.path.join(LOG_FOLDER, fname), 'r', encoding='utf-8') as f:
                    meta = {}
                    for line in f:
                        if line.startswith('# Created:'):
                            meta['created'] = line.strip().split(':',1)[1].strip()
                        elif line.startswith('# Last changed:'):
                            meta['last_changed'] = line.strip().split(':',1)[1].strip()
                        elif not line.startswith('#') and not line.strip()=='':
                            break
                    self.log_meta[proj] = meta if meta else {
                        'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'last_changed': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                with open(os.path.join(LOG_FOLDER, fname), 'r', encoding='utf-8') as f:
                    reader = csv.reader(filter(lambda row: row[0] not in '#', f))
                    next(reader, None) # skip header
                    for row in reader:
                        if len(row) >= 3:
                            self.sessions[proj].append({
                                'start': row[0],'end': row[1],'duration': int(row[2]),'window_title': row[3] if len(row) > 3 else ''
                            })

    def on_window_open(self, window_title):
        proj = extract_project_name(window_title)
        if proj not in self.session_active:
            self.session_active[proj] = {
                'start': datetime.now(),
                'window_title': window_title
            }

    def on_window_close(self, window_title):
        proj = extract_project_name(window_title)
        if proj in self.session_active:
            session = self.session_active.pop(proj)
            end_time = datetime.now()
            duration = int((end_time - session['start']).total_seconds())
            self.sessions[proj].append({
                'start': session['start'].strftime('%Y-%m-%d %H:%M:%S'),
                'end': end_time.strftime('%Y-%m-%d %H:%M:%S'),
                'duration': duration,
                'window_title': window_title
            })
            self.save_session_log(proj)

    def save_session_log(self, proj):
        if not self.sessions[proj]: return
        fname = os.path.join(LOG_FOLDER, f"{proj}_log.csv")
        total_time = sum(s['duration'] for s in self.sessions[proj])
        created = self.log_meta.get(proj,{}).get('created', self.sessions[proj][0]['start'])
        last_changed = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_meta[proj] = {'created': created,'last_changed': last_changed}
        header_lines = [
            f'# Created: {created}',
            f'# Last changed: {last_changed}',
            f'# Total time across all sessions (sec): {total_time}',
            '# Columns: session_start, session_end, session_duration_sec, window_title'
        ]
        with open(fname, 'w', encoding='utf-8', newline='') as f:
            for line in header_lines:
                f.write(line+'\n')
            writer = csv.writer(f)
            writer.writerow(['session_start','session_end','session_duration_sec','window_title'])
            for s in self.sessions[proj]:
                writer.writerow([s['start'], s['end'], s['duration'], s['window_title']])

    def run(self):
        prev_states = {}
        try:
            while True:
                # Poll windows
                current_windows = get_all_open_windows()
                # Track starts
                tracked_now = set()
                for wtitle in current_windows:
                    proj = extract_project_name(wtitle)
                    # Only track selected architecture apps:
                    if get_app_name_from_title(wtitle).lower() in {
                        'blender', 'archicad', 'revit', 'autocad', 'sketchup',
                        'rhino', '3dsmax', 'lumion', 'enscape', 'photoshop',
                        'illustrator', 'vectorworks'
                    }:
                        tracked_now.add(proj)
                        if proj not in self.session_active:
                            self.on_window_open(wtitle)
                # Handle window closes
                closed_projs = set(self.session_active) - tracked_now
                for proj in list(closed_projs):
                    wtitle = self.session_active[proj]['window_title']
                    self.on_window_close(wtitle)
                time.sleep(CHECK_INTERVAL)
        except KeyboardInterrupt:
            # Save any remaining sessions as closed NOW
            for proj in list(self.session_active):
                wtitle = self.session_active[proj]['window_title']
                self.on_window_close(wtitle)
            print("Tracker stopped. Logs saved.")

if __name__ == "__main__":
    tracker = WindowTimeTracker()
    tracker.run()


# ============================================
# PROGRAM START
# ============================================
# This is where the program actually starts running

if __name__ == "__main__":
    # Create the tracker object
    tracker = WindowTimeTracker()
    # Start tracking (runs until you press Ctrl+C)
    tracker.run()
