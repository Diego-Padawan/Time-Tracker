# ============================================
# PROJECT TIME VISUALIZATION TOOL
# ============================================
# This program reads your window tracker logs and creates visual charts

import os
import csv
from datetime import datetime, timedelta
from collections import defaultdict
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Rectangle

# ============================================
# CONFIGURATION
# ============================================
LOG_FOLDER = "window_logs"  # Same folder name as in window_tracker.py

# ============================================
# READ DATA FROM LOG FILES
# ============================================

def read_project_logs():
    """
    Reads all project log files and extracts session data.
    Returns a dictionary with project names and their sessions.
    """
    project_data = {}
    
    if not os.path.exists(LOG_FOLDER):
        print(f"Error: The folder '{LOG_FOLDER}' doesn't exist yet.")
        print("Run window_tracker.py first to create some logs!")
        return project_data
    
    # Go through each log file in the folder
    for filename in os.listdir(LOG_FOLDER):
        if filename.endswith('_log.csv'):
            project_name = filename[:-8]  # Remove '_log.csv' from filename
            filepath = os.path.join(LOG_FOLDER, filename)
            
            sessions = []
            total_seconds = 0
            
            # Read the CSV file
            with open(filepath, 'r', encoding='utf-8') as f:
                # Skip comment lines starting with #
                lines = [line for line in f if not line.startswith('#')]
                
                # Parse CSV data
                reader = csv.DictReader(lines)
                for row in reader:
                    try:
                        start_time = datetime.strptime(row['session_start'], '%Y-%m-%d %H:%M:%S')
                        end_time = datetime.strptime(row['session_end'], '%Y-%m-%d %H:%M:%S')
                        duration = int(row['session_duration_sec'])
                        
                        sessions.append({
                            'start': start_time,
                            'end': end_time,
                            'duration': duration,
                            'date': start_time.date()
                        })
                        total_seconds += duration
                    except Exception as e:
                        # Skip rows with errors
                        continue
            
            if sessions:
                project_data[project_name] = {
                    'sessions': sessions,
                    'total_seconds': total_seconds
                }
    
    return project_data

# ============================================
# CREATE TIMELINE CHART
# ============================================

def create_timeline_chart(project_data):
    """
    Creates a visual timeline showing which days you worked on each project.
    """
    if not project_data:
        print("No data to visualize!")
        return
    
    # Group sessions by date for each project
    project_dates = {}
    for project_name, data in project_data.items():
        dates_with_time = defaultdict(float)  # date -> total hours
        
        for session in data['sessions']:
            date = session['date']
            hours = session['duration'] / 3600  # Convert seconds to hours
            dates_with_time[date] += hours
        
        project_dates[project_name] = dates_with_time
    
    # Create the chart
    fig, ax = plt.subplots(figsize=(14, max(6, len(project_data) * 0.8)))
    
    # Colors for different projects
    colors = plt.cm.tab10(range(len(project_data)))
    
    # Plot each project
    project_names = list(project_dates.keys())
    for idx, project_name in enumerate(project_names):
        dates_dict = project_dates[project_name]
        dates = list(dates_dict.keys())
        hours = [dates_dict[d] for d in dates]
        
        # Plot horizontal bars for each date
        y_position = idx
        ax.barh([y_position] * len(dates), hours, left=mdates.date2num(dates), 
                height=0.6, color=colors[idx], alpha=0.7, label=project_name)
    
    # Format the chart
    ax.set_yticks(range(len(project_names)))
    ax.set_yticklabels(project_names)
    ax.set_xlabel('Date', fontsize=12)
    ax.set_ylabel('Project', fontsize=12)
    ax.set_title('Project Work Timeline', fontsize=14, fontweight='bold')
    
    # Format x-axis to show dates nicely
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
    
    ax.grid(axis='x', alpha=0.3)
    plt.tight_layout()
    
    # Save the chart
    plt.savefig('project_timeline.png', dpi=150, bbox_inches='tight')
    print("✓ Timeline chart saved as 'project_timeline.png'")
    
    plt.show()

# ============================================
# CREATE TOTAL TIME SUMMARY CHART
# ============================================

def create_summary_chart(project_data):
    """
    Creates a bar chart showing total time spent on each project.
    """
    if not project_data:
        return
    
    # Calculate total hours for each project
    project_names = []
    total_hours = []
    
    for project_name, data in project_data.items():
        hours = data['total_seconds'] / 3600
        project_names.append(project_name)
        total_hours.append(hours)
    
    # Sort by total time (most time first)
    sorted_data = sorted(zip(project_names, total_hours), key=lambda x: x[1], reverse=True)
    project_names, total_hours = zip(*sorted_data) if sorted_data else ([], [])
    
    # Create the chart
    fig, ax = plt.subplots(figsize=(12, max(6, len(project_names) * 0.5)))
    
    # Create horizontal bar chart
    colors = plt.cm.viridis(range(len(project_names)))
    bars = ax.barh(project_names, total_hours, color=colors, alpha=0.7)
    
    # Add hour labels on the bars
    for bar, hours in zip(bars, total_hours):
        days = int(hours // 24)
        remaining_hours = hours % 24
        label = f"{int(hours)}h"
        if days > 0:
            label = f"{days}d {int(remaining_hours)}h"
        
        ax.text(bar.get_width() + max(total_hours) * 0.01, bar.get_y() + bar.get_height()/2, 
                label, va='center', fontsize=10, fontweight='bold')
    
    ax.set_xlabel('Total Hours', fontsize=12)
    ax.set_ylabel('Project', fontsize=12)
    ax.set_title('Total Time Spent on Each Project', fontsize=14, fontweight='bold')
    ax.grid(axis='x', alpha=0.3)
    
    plt.tight_layout()
    
    # Save the chart
    plt.savefig('project_summary.png', dpi=150, bbox_inches='tight')
    print("✓ Summary chart saved as 'project_summary.png'")
    
    plt.show()

# ============================================
# PRINT TEXT SUMMARY
# ============================================

def print_text_summary(project_data):
    """
    Prints a text summary of your project time in the console.
    """
    if not project_data:
        print("\nNo project data found yet!")
        return
    
    print("\n" + "="*60)
    print("PROJECT TIME SUMMARY")
    print("="*60)
    
    # Sort projects by total time
    sorted_projects = sorted(project_data.items(), 
                            key=lambda x: x[1]['total_seconds'], 
                            reverse=True)
    
    for project_name, data in sorted_projects:
        total_hours = data['total_seconds'] / 3600
        days = int(total_hours // 24)
        hours = int(total_hours % 24)
        minutes = int((total_hours % 1) * 60)
        
        # Count unique work days
        unique_dates = set(session['date'] for session in data['sessions'])
        num_work_days = len(unique_dates)
        
        print(f"\n📁 {project_name}")
        print(f"   Total time: {days} days, {hours} hours, {minutes} minutes")
        print(f"   Work days: {num_work_days} days")
        print(f"   Sessions: {len(data['sessions'])}")
    
    print("\n" + "="*60)

# ============================================
# MAIN PROGRAM
# ============================================

def main():
    print("="*60)
    print("PROJECT TIME VISUALIZATION TOOL")
    print("="*60)
    print("\nReading your project logs...")
    
    # Read all project data
    project_data = read_project_logs()
    
    if not project_data:
        print("\nNo log files found! Run window_tracker.py first.")
        return
    
    print(f"\nFound {len(project_data)} projects!")
    
    # Show text summary
    print_text_summary(project_data)
    
    # Create visualizations
    print("\nCreating visualizations...")
    create_timeline_chart(project_data)
    create_summary_chart(project_data)
    
    print("\n✓ Done! Check the PNG files in this folder.")
    print("  - project_timeline.png: Shows when you worked on each project")
    print("  - project_summary.png: Shows total time for each project")

if __name__ == "__main__":
    main()
