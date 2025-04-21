# Required installations:
# pip install sounddevice numpy SpeechRecognition pyttsx3 openai pillow pystray pywin32

import sys
import threading
import sounddevice as sd
import numpy as np
import speech_recognition as sr
import pyttsx3
import tkinter as tk
from tkinter import scrolledtext, ttk, Menu, Toplevel, Label, Entry, Button, Frame, Canvas
import sqlite3
import openai
import time
import os
from PIL import Image, ImageTk, ImageSequence
import traceback
import pystray
from pystray import Icon, MenuItem as item
import ctypes
from ctypes import wintypes
from urllib.request import urlopen
import socket
from itertools import cycle
import queue
import win32gui
import win32con
import winshell
import pythoncom
from win32com.client import Dispatch
import shutil  # For copying files

# Global variables
tray_icon = None
tray_icon_initialized = False
logo_photo = None
computer_photo = None
loading_label = None
loading_frames = None
loading_thread = None
message_queue = queue.Queue()

# Initialize root window first
root = tk.Tk()
root.title('🎧 Speech → 🧠 AI Grammar Fix → 🗣️ Voice')
root.geometry('700x650')

# Paths
base_dir = os.path.dirname(__file__)
DB_PATH = os.path.join(base_dir, 'data.sqlite')
app_icon_path = os.path.join(base_dir, 'needyamin.ico')
logo_display_path = os.path.join(base_dir, 'logo_display.png')
logo_comp_path = os.path.join(base_dir, 'logo_display_computer.png')

# Set window icon (both titlebar and taskbar)
if os.path.exists(app_icon_path):
    try:
        # Set the taskbar icon
        myappid = 'needyamin.speechapp.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        root.iconbitmap(default=app_icon_path)
        root.iconbitmap(app_icon_path)
    except Exception as e:
        print(f"Error setting icon: {e}")

# Load images after root window creation
if os.path.exists(logo_display_path):
    img = Image.open(logo_display_path).resize((16,16), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(img)
if os.path.exists(logo_comp_path):
    img2 = Image.open(logo_comp_path).resize((16,16), Image.LANCZOS)
    computer_photo = ImageTk.PhotoImage(img2)

# Database setup
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()
cur.execute('''CREATE TABLE IF NOT EXISTS config (
    id INTEGER PRIMARY KEY,
    api_key TEXT,
    api_base TEXT,
    password TEXT
)''')
conn.commit()
cur.execute('SELECT api_key, api_base FROM config WHERE id=1')
row = cur.fetchone()
if row:
    openai.api_key, openai.api_base = row
else:
    openai.api_key = ''
    openai.api_base = ''
    cur.execute('INSERT OR REPLACE INTO config (id, api_key, api_base, password) VALUES (1, ?, ?, ?)',
                ('', '', 'admin'))
    conn.commit()

# Initialize recognizer and TTS
recognizer = sr.Recognizer()
tts = pyttsx3.init()
recording = threading.Event()
frames = []
stream = None

# Create loading GIF function
def create_loading_gif():
    frames = []
    size = (30, 30)
    for i in range(8):
        frame = Image.new('RGBA', size, (255, 255, 255, 0))
        draw = ImageDraw.Draw(frame)
        angle = i * 45
        draw.arc([2, 2, 27, 27], angle, angle + 300, fill='blue', width=2)
        frames.append(frame)
    
    gif_path = os.path.join(base_dir, 'load.gif')
    frames[0].save(gif_path, save_all=True, append_images=frames[1:],
                   duration=100, loop=0)

def check_internet():
    try:
        urlopen('http://www.google.com', timeout=1)
        return True
    except:
        return False

def show_no_internet_dialog():
    dialog = Toplevel(root)
    dialog.title('No Internet Connection')
    dialog.geometry('300x150')
    dialog.transient(root)
    dialog.grab_set()
    
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f'{width}x{height}+{x}+{y}')
    
    Label(dialog, text='No Internet Connection', font=('Arial', 12, 'bold')).pack(pady=10)
    Label(dialog, text='Please check your internet connection\nand try again.').pack(pady=10)
    Button(dialog, text='Retry', command=lambda: retry_connection(dialog)).pack(pady=10)
    Button(dialog, text='Exit', command=root.quit).pack(pady=5)

def retry_connection(dialog):
    if check_internet():
        dialog.destroy()
        log('Internet connection restored.', level='INFO')
    else:
        log('Still no internet connection.', level='ERROR')

def create_rounded_frame(parent, bg_color, padding=10):
    frame = Frame(parent, bg=parent['bg'])
    
    canvas = Canvas(frame, 
                   bg=parent['bg'],
                   highlightthickness=0,
                   width=500,
                   height=100)
    canvas.pack(fill='both', expand=True)
    
    # Create rounded rectangle
    radius = 15
    points = [
        radius, 0,
        500-radius, 0,
        500, radius,
        500, 100-radius,
        500-radius, 100,
        radius, 100,
        0, 100-radius,
        0, radius
    ]
    
    canvas.create_polygon(points, 
                         smooth=True,
                         fill=bg_color,
                         outline=bg_color)
    
    # Create inner frame for content
    inner_frame = Frame(canvas, bg=bg_color)
    canvas.create_window(padding, padding, 
                        window=inner_frame,
                        anchor='nw')
    
    return frame, inner_frame

def create_message_bubble(parent, message, is_bot=True, icon_image=None):
    frame = Frame(parent, bg=parent['bg'])
    
    bubble_color = '#E8E8E8' if is_bot else '#DCF8C6'
    bubble_frame, content_frame = create_rounded_frame(frame, bubble_color)
    bubble_frame.pack(pady=5, padx=10, anchor='w' if is_bot else 'e')
    
    # Time label
    time_label = Label(content_frame, 
                      text=time.strftime('%d %B %Y: %I:%M %p').lstrip('0'),
                      font=('Consolas', 8),
                      fg='gray',
                      bg=bubble_color)
    time_label.pack(padx=5, pady=(5,0), anchor='w')
    
    # Message content frame
    msg_content = Frame(content_frame, bg=bubble_color)
    msg_content.pack(fill='x', padx=5, pady=5)
    
    # Icon (if provided)
    if icon_image:
        icon_label = Label(msg_content, image=icon_image, bg=bubble_color)
        icon_label.pack(side='left', padx=(0,5))
    
    # Message text
    msg_label = Label(msg_content,
                     text=message,
                     wraplength=400,
                     justify='left',
                     bg=bubble_color,
                     font=('Consolas', 11))
    msg_label.pack(side='left', fill='x', expand=True)
    
    return frame
	
def log(message, level='INFO', icon_image=None):
    output.configure(state='normal')
    
    if message == 'Clear Logs':
        output.delete('1.0', 'end')
        output.configure(state='disabled')
        return
    
    # Determine if this is a bot message or user message
    is_bot = message.startswith('Corrected:')
    
    # Clean up the message
    if message.startswith('Corrected:'):
        message = message[10:].strip()
    elif message.startswith('Recognized text:'):
        message = message[15:].strip()
    
    # Create and pack the message bubble
    bubble = create_message_bubble(output, message, is_bot, icon_image)
    output.window_create('end', window=bubble)
    output.insert('end', '\n')
    
    output.configure(state='disabled')
    output.see('end')

def log_plain(message, level='INFO'):
    """Log plain messages without creating a conversation bubble."""
    output.configure(state='normal')
    timestamp = time.strftime('%d %B %Y: %I:%M %p').lstrip('0')
    output.insert('end', f'{timestamp} - {message}\n', level)
    output.configure(state='disabled')
    output.see('end')

def load_animation():
    global loading_frames
    try:
        gif_path = os.path.join(base_dir, 'load.gif')
        if not os.path.exists(gif_path):
            create_loading_gif()
        gif = Image.open(gif_path)
        frames = []
        for frame in ImageSequence.Iterator(gif):
            frame = frame.copy()
            frame = frame.resize((30, 30), Image.LANCZOS)
            frames.append(ImageTk.PhotoImage(frame))
        loading_frames = cycle(frames)
        return True
    except Exception as e:
        log(f'Error loading animation: {e}', level='ERROR')
        return False

def update_loading_animation():
    global loading_label
    if loading_label and loading_frames:
        try:
            frame = next(loading_frames)
            loading_label.configure(image=frame)
            loading_label.image = frame
            root.after(50, update_loading_animation)
        except Exception as e:
            log(f'Error updating animation: {e}', level='ERROR')

def show_loading():
    global loading_label
    if not loading_label and loading_frames:
        loading_label = Label(root)
        loading_label.place(relx=0.5, rely=0.5, anchor='center')
        update_loading_animation()

def hide_loading():
    global loading_label
    if loading_label:
        loading_label.place_forget()
        loading_label = None

def verify_password():
    def login():
        pwd = pass_entry.get().strip()
        cur.execute('SELECT password FROM config WHERE id=1')
        if pwd == cur.fetchone()[0]:
            pass_win.destroy()
            open_settings_interface()
        else:
            log('Incorrect password.', level='ERROR')
    pass_win = Toplevel(root)
    pass_win.title('Login')
    Label(pass_win, text='Enter Password:').grid(row=0, column=0, padx=5, pady=5)
    pass_entry = Entry(pass_win, show='*')
    pass_entry.grid(row=0, column=1, padx=5, pady=5)
    Button(pass_win, text='Login', command=login).grid(row=1, column=0, columnspan=2, pady=10)

def open_settings_interface():
    def save():
        key = key_entry.get().strip()
        base = base_entry.get().strip()
        newp = new_entry.get().strip()
        conf = conf_entry.get().strip()
        cur.execute('SELECT password FROM config WHERE id=1')
        oldp = cur.fetchone()[0]
        if newp and newp != conf:
            log('Passwords do not match.', level='ERROR')
            return
        cur.execute('UPDATE config SET api_key=?, api_base=?, password=? WHERE id=1',
                    (key, base, newp or oldp))
        conn.commit()
        openai.api_key, openai.api_base = key, base
        log('Settings saved.')
        sett.destroy()
    
    sett = Toplevel(root)
    sett.title('Settings')
    Label(sett, text='OpenAI API Key:').grid(row=0, column=0, padx=5, pady=5)
    key_entry = Entry(sett, width=50)
    key_entry.grid(row=0, column=1, padx=5, pady=5)
    key_entry.insert(0, openai.api_key)
    Label(sett, text='API Base URL:').grid(row=1, column=0, padx=5, pady=5)
    base_entry = Entry(sett, width=50)
    base_entry.grid(row=1, column=1, padx=5, pady=5)
    base_entry.insert(0, openai.api_base)
    Label(sett, text='New Password:').grid(row=2, column=0, padx=5, pady=5)
    new_entry = Entry(sett, show='*', width=50)
    new_entry.grid(row=2, column=1, padx=5, pady=5)
    Label(sett, text='Confirm Password:').grid(row=3, column=0, padx=5, pady=5)
    conf_entry = Entry(sett, show='*', width=50)
    conf_entry.grid(row=3, column=1, padx=5, pady=5)
    Button(sett, text='Save', command=save).grid(row=4, column=0, columnspan=2, pady=10)

def open_settings():
    verify_password()

def audio_callback(indata, frames_count, time_info, status):
    if status:
        log(f'Audio status: {status}', level='WARNING')
    frames.append(indata.copy())

# Add a status bar at the bottom of the GUI
status_label = Label(root, text="Ready", anchor="w", font=("Arial", 10), bg="#f0f0f0", relief="sunken")
status_label.grid(row=2, column=0, columnspan=5, sticky="ew", padx=5, pady=5)

def update_status(message):
    """Update the status bar with the given message."""
    status_label.config(text=message)

def start_recording():
    global stream
    frames.clear()
    recording.set()
    start_btn.config(state='disabled')
    stop_btn.config(state='normal')
    update_status("Recording started... Click 'Stop' when done.")  # Update the status bar

    def run_stream():
        global stream
        try:
            stream = sd.InputStream(samplerate=16000, channels=1, callback=audio_callback)
            with stream:
                while recording.is_set():
                    sd.sleep(100)
        except Exception as e:
            log(f'Error starting recording: {e}', level='ERROR')
        finally:
            stream = None
            stop_btn.config(state='disabled')
            start_btn.config(state='normal')

    threading.Thread(target=run_stream, daemon=True).start()

def stop_recording():
    recording.clear()
    update_status("Recording stopped.")  # Update the status bar
    threading.Thread(target=process_audio, daemon=True).start()

def process_audio():
    if not frames:
        log('No audio captured.', level='WARNING')
        return
    audio_np = np.concatenate(frames, axis=0).astype(np.float32)
    if np.max(np.abs(audio_np)) == 0:
        log('Silence detected.', level='WARNING')
        return
    audio_np = (audio_np / np.max(np.abs(audio_np)) * 32767).astype(np.int16)
    audio_bytes = audio_np.tobytes()
    audio_data = sr.AudioData(audio_bytes, 16000, 2)
    try:
        text = recognizer.recognize_google(audio_data, language='en-US')
    except sr.UnknownValueError:
        log('Could not understand audio.', level='WARNING')
        return
    except sr.RequestError as e:
        log(f'Recognition API error: {e}', level='ERROR')
        return
    show_correction(text)

def show_correction(text):
    if not check_internet():
        show_no_internet_dialog()
        return

    log(text, icon_image=logo_photo)  # Display the user's input text
    show_loading()  # Show the loading animation

    def process_correction():
        try:
            corrected = correct_grammar_gpt(text)
            message_queue.put(('success', corrected))
        except Exception as e:
            message_queue.put(('error', str(e)))

    def check_result():
        try:
            if not message_queue.empty():
                status, result = message_queue.get()
                hide_loading()  # Hide the loading animation

                if status == 'success':
                    log(result, icon_image=computer_photo)  # Display the corrected message
                    root.after(100, lambda: play_voice(result))  # Start voice output after displaying the message
                else:
                    log(f'Error: {result}', level='ERROR')
                return

            root.after(100, check_result)  # Check again after 100ms
        except Exception as e:
            hide_loading()
            log(f'Error: {e}', level='ERROR')

    threading.Thread(target=process_correction, daemon=True).start()
    root.after(100, check_result)

def play_voice(text):
    """Play the corrected text using TTS."""
    try:
        tts.say(text)
        tts.runAndWait()
    except Exception as e:
        log(f'Error during voice playback: {e}', level='ERROR')

def correct_grammar_gpt(text):
    try:
        resp = openai.ChatCompletion.create(
            model='gpt-3.5-turbo',
            messages=[
                {'role':'system','content':'You are an English grammar assistant.'},
                {'role':'user','content':text}
            ]
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        return f'(Error: {e})'

# GUI Setup
menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label='Settings', command=open_settings)
file_menu.add_command(label='Clear Logs', command=lambda: log('Clear Logs'))
file_menu.add_separator()
file_menu.add_command(label='Exit', command=root.quit)
menu_bar.add_cascade(label='File', menu=file_menu)
help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label='About', command=lambda: log('Speech-to-Text UI v1.0'))
menu_bar.add_cascade(label='Help', menu=help_menu)
root.config(menu=menu_bar)

output = scrolledtext.ScrolledText(root, 
                                 wrap=tk.WORD, 
                                 state='disabled', 
                                 font=('Consolas', 11), 
                                 bg='#f0f0f0',
                                 padx=10,
                                 pady=10)
output.grid(row=0, column=0, columnspan=5, padx=10, pady=10, sticky='nsew')

# Add style configurations
style = ttk.Style()
style.configure('Chat.TFrame', background='#f0f0f0')

btn_frame = tk.Frame(root)
btn_frame.grid(row=1, column=0, columnspan=5, pady=5, sticky='ew')
start_btn = Button(btn_frame, text='Start Recording', command=start_recording)
stop_btn = Button(btn_frame, text='Stop Recording', state='disabled', command=stop_recording)
clear_btn = Button(btn_frame, text='Clear Logs', command=lambda: log('Clear Logs'))
settings_btn = Button(btn_frame, text='Settings', command=open_settings)
start_btn.grid(row=0, column=0, padx=5)
stop_btn.grid(row=0, column=1, padx=5)
clear_btn.grid(row=0, column=2, padx=5)
settings_btn.grid(row=0, column=3, padx=5)

# Configure microphone dropdown
mic_devices = [d['name'] for d in sd.query_devices() if d['max_input_channels']>0]
mic_dropdown = ttk.Combobox(btn_frame, values=mic_devices)
if mic_devices:
    mic_dropdown.set(mic_devices[0])
mic_dropdown.grid(row=0, column=4, padx=5, sticky='ew')

# Configure grid weights
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Tray integration
def restore(icon, item):
    global tray_icon_initialized
    icon.stop()
    tray_icon_initialized = False
    root.deiconify()

def quit_app(icon, item):
    icon.stop()
    root.destroy()

def setup_tray():
    global tray_icon, tray_icon_initialized
    if tray_icon_initialized:
        return True
    img = Image.open(app_icon_path) if os.path.exists(app_icon_path) else None
    tray_icon = Icon('SpeechApp', img, 'Speech Assistant', menu=(item('Restore', restore), item('Quit', quit_app)))
    tray_icon_initialized = True
    threading.Thread(target=tray_icon.run, daemon=True).start()
    return True

# Center the window on screen
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

# Bind events
root.bind('<Unmap>', lambda e: (root.state()=='iconic' and setup_tray() and root.withdraw()))
root.protocol('WM_DELETE_WINDOW', lambda: (root.withdraw(), setup_tray()))

# Center the main window
center_window(root)

# Initialize loading animation
if not load_animation():
    log('Warning: Loading animation not available.', level='WARNING')

# Check internet connection
if not check_internet():
    root.after(1000, show_no_internet_dialog)

# Paths for auto-copy and startup
ansnewtech_dir = r"C:\ANSNEWTECH"
startup_folder = winshell.startup()
startup_shortcut_path = os.path.join(startup_folder, "Speech Assistant.lnk")
start_menu_shortcut_path = os.path.join(winshell.start_menu(), "Speech Assistant.lnk")

# Copy the entire program to C:\ANSNEWTECH and run from there
def copy_to_ansnewtech():
    """Copy the entire program to C:\ANSNEWTECH and restart from there."""
    try:
        if not os.path.exists(ansnewtech_dir):
            os.makedirs(ansnewtech_dir)  # Create the directory if it doesn't exist

        # Copy all files in the current directory to C:\ANSNEWTECH
        for filename in os.listdir(base_dir):
            src_path = os.path.join(base_dir, filename)
            dest_path = os.path.join(ansnewtech_dir, filename)
            if os.path.isfile(src_path):
                shutil.copy2(src_path, dest_path)

        # Restart the program from C:\ANSNEWTECH
        target_script_path = os.path.join(ansnewtech_dir, os.path.basename(__file__))
        os.execl(sys.executable, sys.executable, target_script_path)
    except Exception as e:
        print(f"Error copying program to {ansnewtech_dir}: {e}")

# Ensure the program runs from C:\ANSNEWTECH
if __file__.lower() != os.path.join(ansnewtech_dir, os.path.basename(__file__)).lower():
    copy_to_ansnewtech()

# Create a shortcut in the Windows Startup folder
def create_startup_shortcut():
    """Create a shortcut for the app in the Windows Startup folder."""
    try:
        target = sys.executable  # Python executable
        target_script_path = os.path.join(ansnewtech_dir, os.path.basename(__file__))

        # Create the shortcut
        shell = Dispatch("WScript.Shell", pythoncom.CoInitialize())
        shortcut = shell.CreateShortcut(startup_shortcut_path)
        shortcut.TargetPath = target
        shortcut.Arguments = f'"{target_script_path}"'
        shortcut.WorkingDirectory = ansnewtech_dir
        shortcut.IconLocation = os.path.join(ansnewtech_dir, "needyamin.ico")
        shortcut.Save()

        print(f"Startup shortcut created at: {startup_shortcut_path}")
    except Exception as e:
        print(f"Error creating startup shortcut: {e}")

# Create a shortcut in the Windows Start Menu
def create_start_menu_shortcut():
    """Create a shortcut for the app in the Windows Start Menu."""
    try:
        target = sys.executable  # Python executable
        target_script_path = os.path.join(ansnewtech_dir, os.path.basename(__file__))

        # Create the shortcut
        shell = Dispatch("WScript.Shell", pythoncom.CoInitialize())
        shortcut = shell.CreateShortcut(start_menu_shortcut_path)
        shortcut.TargetPath = target
        shortcut.Arguments = f'"{target_script_path}"'
        shortcut.WorkingDirectory = ansnewtech_dir
        shortcut.IconLocation = os.path.join(ansnewtech_dir, "needyamin.ico")
        shortcut.Save()

        print(f"Start Menu shortcut created at: {start_menu_shortcut_path}")
    except Exception as e:
        print(f"Error creating Start Menu shortcut: {e}")

# Remove the startup shortcut
def remove_startup_shortcut():
    """Remove the startup shortcut."""
    try:
        if os.path.exists(startup_shortcut_path):
            os.remove(startup_shortcut_path)
            print(f"Startup shortcut removed: {startup_shortcut_path}")
    except Exception as e:
        print(f"Error removing startup shortcut: {e}")

# Check if the startup shortcut exists
def is_startup_enabled():
    """Check if the startup shortcut exists."""
    return os.path.exists(startup_shortcut_path)

# Add "Start on Startup" and "Stop on Startup" options to the Help menu with tick icons
def update_startup_menu():
    """Update the Help menu to show the current startup status with tick icons."""
    help_menu.delete(0, "end")  # Clear existing menu items
    if is_startup_enabled():
        help_menu.add_command(label="✓ Start on Startup", command=lambda: toggle_startup(True))
        help_menu.add_command(label="Stop on Startup", command=lambda: toggle_startup(False))
    else:
        help_menu.add_command(label="Start on Startup", command=lambda: toggle_startup(True))
        help_menu.add_command(label="✓ Stop on Startup", command=lambda: toggle_startup(False))
    help_menu.add_separator()
    help_menu.add_command(label="About", command=lambda: log("Speech-to-Text UI v1.0"))

def toggle_startup(enable):
    """Enable or disable startup functionality."""
    if enable:
        create_startup_shortcut()
        log("Startup enabled. The program will start automatically on system boot.", level="INFO")
    else:
        remove_startup_shortcut()
        log("Startup disabled. The program will no longer start automatically on system boot.", level="INFO")
    update_startup_menu()  # Refresh the Help menu

# Add startup menu options
update_startup_menu()

# Create all shortcuts
create_startup_shortcut()
create_start_menu_shortcut()

# Start the application
root.mainloop()