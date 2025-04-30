# Required installations:
# pip install sounddevice numpy SpeechRecognition pyttsx3 openai pillow pystray pywin32

import sys
import threading
import sounddevice as sd
import numpy as np
import speech_recognition as sr
import pyttsx3
import tkinter as tk
from tkinter import scrolledtext, ttk, Menu, Toplevel, Label, Entry, Button, Frame, Canvas, Checkbutton, BooleanVar, Text
import sqlite3
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
import shutil
import re
import openai

# Global variables
tray_icon = None
tray_icon_initialized = False
logo_photo = None
computer_photo = None
loading_label = None
loading_frames = None
loading_thread = None
message_queue = queue.Queue()
is_speaking = False
tts_engine = None

# Initialize root window first
root = tk.Tk()
root.title('🎧 Speech → 🧠 AI Grammar Fix → 🗣️ Voice')
root.geometry('800x700')  # Increased window size

# Set modern theme colors
BG_COLOR = '#f5f5f5'
PRIMARY_COLOR = '#2196F3'
SECONDARY_COLOR = '#1976D2'
ACCENT_COLOR = '#FF4081'
TEXT_COLOR = '#333333'
SUCCESS_COLOR = '#4CAF50'
WARNING_COLOR = '#FFC107'
ERROR_COLOR = '#F44336'

# Configure root window
root.configure(bg=BG_COLOR)
root.option_add('*Font', 'Arial 10')

# Configure ttk styles
style = ttk.Style()
style.theme_use('clam')  # Use clam theme as base

# Configure button styles
style.configure('Modern.TButton',
    background=PRIMARY_COLOR,
    foreground='white',
    padding=10,
    relief='flat',
    font=('Arial', 10, 'bold')
)
style.map('Modern.TButton',
    background=[('active', SECONDARY_COLOR)],
    foreground=[('active', 'white')]
)

# Configure entry styles
style.configure('Modern.TEntry',
    fieldbackground='white',
    foreground=TEXT_COLOR,
    padding=5,
    relief='flat'
)

# Configure label styles
style.configure('Modern.TLabel',
    background=BG_COLOR,
    foreground=TEXT_COLOR,
    font=('Arial', 10)
)

# Configure combobox styles
style.configure('Modern.TCombobox',
    fieldbackground='white',
    foreground=TEXT_COLOR,
    padding=5,
    relief='flat'
)

# Configure scrollbar styles
style.configure('Modern.Vertical.TScrollbar',
    background=BG_COLOR,
    troughcolor=BG_COLOR,
    relief='flat'
)

# Initialize Tkinter variables after root window
use_online_mode = BooleanVar(root, value=False)

# Paths
base_dir = os.path.dirname(__file__)
DB_PATH = os.path.join(base_dir, 'data.sqlite')
app_icon_path = os.path.join(base_dir, 'needyamin.ico')
logo_display_path = os.path.join(base_dir, 'logo_display.png')
logo_comp_path = os.path.join(base_dir, 'logo_display_computer.png')

# Set permanent API base URL
OPENAI_API_BASE = "https://openrouter.ai/api/v1"

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
    img = Image.open(logo_display_path).resize((32,32), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(img)
if os.path.exists(logo_comp_path):
    img2 = Image.open(logo_comp_path).resize((32,32), Image.LANCZOS)
    computer_photo = ImageTk.PhotoImage(img2)

# Database setup
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()
cur.execute('''CREATE TABLE IF NOT EXISTS config (
    id INTEGER PRIMARY KEY,
    api_key TEXT,
    password TEXT
)''')
conn.commit()
cur.execute('SELECT api_key FROM config WHERE id=1')
row = cur.fetchone()
if row:
    openai.api_key = row[0]
else:
    openai.api_key = ''
    cur.execute('INSERT OR REPLACE INTO config (id, api_key, password) VALUES (1, ?, ?)',
                ('', 'admin'))
    conn.commit()

# Set the permanent API base URL
openai.api_base = OPENAI_API_BASE

# Load saved settings
cur.execute('SELECT api_key, password FROM config WHERE id=1')
row = cur.fetchone()
if row:
    openai.api_key = row[0]
    use_online_mode.set(row[1] != '')
else:
    openai.api_key = ''
    use_online_mode.set(False)
    cur.execute('INSERT INTO config (id, api_key, password) VALUES (1, ?, ?)',
                (openai.api_key, use_online_mode.get()))
    conn.commit()

# Initialize recognizer and TTS
recognizer = sr.Recognizer()
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

def create_message_bubble(parent, message, is_bot=True, icon_image=None):
    frame = Frame(parent, bg=parent['bg'])
    
    # Modern bubble colors
    bubble_color = '#E3F2FD' if is_bot else '#E8F5E9'  # Light blue for bot, light green for user
    hover_color = '#BBDEFB' if is_bot else '#C8E6C9'  # Darker shade for hover
    text_color = TEXT_COLOR
    time_color = '#757575'
    
    bubble_frame = Frame(frame, bg=bubble_color)
    bubble_frame.pack(pady=5, padx=10, anchor='w' if is_bot else 'e', fill='x')
    
    # Time and copy button frame with modern styling
    time_frame = Frame(bubble_frame, bg=bubble_color)
    time_frame.pack(fill='x', padx=5, pady=(1,0))
    
    # Time label with modern styling
    time_label = Label(time_frame, 
                      text=time.strftime('%d %B %Y: %I:%M %p').lstrip('0'),
                      font=('Arial', 8),
                      fg=time_color,
                      bg=bubble_color)
    time_label.pack(side='left', padx=(0,10))
    
    # Modern copy button
    copy_button = Button(time_frame,
                        text='📋',
                        font=('Arial', 8),
                        bg=bubble_color,
                        fg=time_color,
                        relief='flat',
                        cursor='hand2',
                        width=2,
                        command=lambda: copy_text())
    copy_button.pack(side='left')
    
    # Create a frame for the icon at the top
    icon_frame = Frame(bubble_frame, bg=bubble_color)
    icon_frame.pack(fill='x', padx=5, pady=(5,0))
    
    # Icon (if provided) - Always at the top
    if icon_image:
        icon_label = Label(icon_frame, image=icon_image, bg=bubble_color)
        icon_label.pack(side='left')
    
    # Message text frame with padding
    text_frame = Frame(bubble_frame, bg=bubble_color)
    text_frame.pack(fill='x', padx=5, pady=(10,5))
    
    # Create a text widget for the message with modern styling
    msg_text = Text(text_frame,
                   wrap='word',
                   width=60,
                   height=1,
                   font=('Arial', 11),
                   bg=bubble_color,
                   fg=text_color,
                   relief='flat',
                   padx=5,
                   pady=0,
                   spacing1=0,
                   spacing2=0,
                   spacing3=0,
                   highlightthickness=0)
    msg_text.pack(side='left', fill='x', expand=True)
    msg_text.insert('1.0', message)
    msg_text.configure(state='disabled')
    
    # Add hover effect to the bubble
    def on_enter(e):
        bubble_frame.configure(bg=hover_color)
        time_frame.configure(bg=hover_color)
        icon_frame.configure(bg=hover_color)
        text_frame.configure(bg=hover_color)
        msg_text.configure(bg=hover_color)
        copy_button.configure(bg=hover_color)
        if icon_image:
            icon_label.configure(bg=hover_color)
    
    def on_leave(e):
        bubble_frame.configure(bg=bubble_color)
        time_frame.configure(bg=bubble_color)
        icon_frame.configure(bg=bubble_color)
        text_frame.configure(bg=bubble_color)
        msg_text.configure(bg=bubble_color)
        copy_button.configure(bg=bubble_color)
        if icon_image:
            icon_label.configure(bg=bubble_color)
    
    bubble_frame.bind('<Enter>', on_enter)
    bubble_frame.bind('<Leave>', on_leave)
    
    # Adjust height based on content
    def adjust_height():
        # Get the text content
        text_content = msg_text.get('1.0', 'end-1c')
        
        # Calculate the number of lines needed
        text_width = msg_text.winfo_width() // 10  # Approximate width in characters
        
        # Split the text into words
        words = text_content.split()
        current_line = []
        lines = []
        
        # Create lines based on the widget width
        for word in words:
            current_line.append(word)
            line_text = ' '.join(current_line)
            if len(line_text) > text_width:
                if len(current_line) > 1:
                    current_line.pop()  # Remove the last word
                    lines.append(' '.join(current_line))
                    current_line = [word]  # Start new line with the word that didn't fit
                else:
                    lines.append(line_text)
                    current_line = []
        
        # Add any remaining words
        if current_line:
            lines.append(' '.join(current_line))
        
        # Calculate the actual height needed
        line_count = len(lines)
        if line_count == 0:
            line_count = 1  # Ensure at least one line for empty messages
        
        # Set the height to match the content
        msg_text.configure(height=line_count)
        
        # Force the text widget to update its height
        msg_text.update_idletasks()
        
        # Get the actual height of the text widget
        actual_height = msg_text.winfo_height()
        
        # Ensure minimum height for short messages
        min_height = 80  # Minimum height in pixels (including logo)
        frame.update_idletasks()
    
    # Call adjust_height after the widget is fully created
    msg_text.after(10, adjust_height)
    
    def copy_text():
        try:
            # Copy text to clipboard
            root.clipboard_clear()
            root.clipboard_append(message)
            root.update()
            
            # Show a temporary "Copied!" label with animation
            copied_label = Label(time_frame,
                               text='✓',
                               font=('Arial', 8),
                               fg=SUCCESS_COLOR,
                               bg=bubble_color)
            copied_label.pack(side='left', padx=(5,0))
            
            # Animate the copied label
            def fade_out():
                copied_label.configure(fg=SUCCESS_COLOR)
                time_frame.after(500, lambda: copied_label.configure(fg=time_color))
                time_frame.after(1000, copied_label.destroy)
            
            fade_out()
        except Exception as e:
            print(f"Error copying text: {e}")
    
    # Force update to ensure button is visible
    frame.update_idletasks()
    
    return frame

def create_rounded_frame(parent, bg_color, padding=10):
    frame = Frame(parent, bg=parent['bg'])
    
    canvas = Canvas(frame, 
                   bg=parent['bg'],
                   highlightthickness=0,
                   width=500,
                   height=100)
    canvas.pack(fill='both', expand=True)
    
    # Create rounded rectangle with modern styling
    radius = 20  # Increased radius for more modern look
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
    
    # Create the main bubble shape
    bubble = canvas.create_polygon(points, 
                                 smooth=True,
                                 fill=bg_color,
                                 outline=bg_color)
    
    # Add subtle shadow effect with a valid color
    shadow = canvas.create_polygon(points,
                                 smooth=True,
                                 fill='#E0E0E0',  # Light gray shadow
                                 outline='#E0E0E0')
    canvas.tag_lower(shadow)
    
    # Create inner frame for content
    inner_frame = Frame(canvas, bg=bg_color)
    canvas.create_window(padding, padding, 
                        window=inner_frame,
                        anchor='nw',
                        width=480)  # Set a fixed width for the inner frame
    
    return frame, inner_frame

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
    
    # Insert at the beginning of the text widget
    output.window_create('1.0', window=bubble)
    output.insert('1.0', '\n')
    
    output.configure(state='disabled')
    # Ensure the latest message is visible
    output.see('1.0')
    
    # Update the message bubble's width after it's been added to the window
    def update_width():
        try:
            for child in bubble.winfo_children():
                if isinstance(child, Frame):
                    for grandchild in child.winfo_children():
                        if isinstance(grandchild, Frame):
                            for widget in grandchild.winfo_children():
                                if isinstance(widget, Text):
                                    window_width = output.winfo_width()
                                    widget.configure(width=min(40, (window_width - 100) // 10))
        except Exception as e:
            print(f"Error updating width: {e}")
    
    # Schedule the update after the window has been drawn
    root.after(100, update_width)

def log_plain(message, level='INFO'):
    """Log plain messages without creating a conversation bubble."""
    output.configure(state='normal')
    timestamp = time.strftime('%d %B %Y: %I:%M %p').lstrip('0')
    # Insert at the beginning
    output.insert('1.0', f'{timestamp} - {message}\n', level)
    output.configure(state='disabled')
    # Ensure the latest message is visible
    output.see('1.0')

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
        newp = new_entry.get().strip()
        conf = conf_entry.get().strip()
        
        # Get current password
        cur.execute('SELECT password FROM config WHERE id=1')
        current_password = cur.fetchone()[0]
        
        if newp and newp != conf:
            log('Passwords do not match.', level='ERROR')
            return
        
        # Save settings
        cur.execute('UPDATE config SET api_key=?, password=? WHERE id=1',
                    (key, newp or current_password))
        conn.commit()
        
        # Update OpenAI configuration
        openai.api_key = key
        
        # Test the API connection
        try:
            client = openai.OpenAI(
                api_key=key,
                base_url=OPENAI_API_BASE
            )
            # Make a test call
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are an English grammar expert. Correct any grammar, spelling, or punctuation errors in the text while maintaining the original meaning. Keep the response concise and clear."},
                    {"role": "user", "content": "Test"}
                ],
                temperature=0.3,
                max_tokens=5
            )
            log('API connection test successful.', level='INFO')
        except Exception as e:
            log(f'API connection test failed: {str(e)}', level='ERROR')
        
        log('Settings saved.')
        sett.destroy()
    
    sett = Toplevel(root)
    sett.title('Settings')
    sett.geometry('500x400')
    sett.resizable(False, False)
    sett.configure(bg=BG_COLOR)
    
    # Center the window
    sett.update_idletasks()
    width = sett.winfo_width()
    height = sett.winfo_height()
    x = (sett.winfo_screenwidth() // 2) - (width // 2)
    y = (sett.winfo_screenheight() // 2) - (height // 2)
    sett.geometry(f'{width}x{height}+{x}+{y}')
    
    # Create main frame
    main_frame = Frame(sett, bg=BG_COLOR, padx=30, pady=30)
    main_frame.pack(fill='both', expand=True)
    
    # Add title
    title_label = Label(main_frame,
                       text="⚙️ Settings",
                       font=('Arial', 16, 'bold'),
                       bg=BG_COLOR,
                       fg=PRIMARY_COLOR)
    title_label.pack(pady=(0, 20))
    
    # Create settings container
    settings_container = Frame(main_frame, bg=BG_COLOR)
    settings_container.pack(fill='both', expand=True)
    
    # Get current settings
    cur.execute('SELECT api_key, password FROM config WHERE id=1')
    current_settings = cur.fetchone()
    current_key = current_settings[0] if current_settings else ''
    
    # API Key section
    api_frame = Frame(settings_container, bg=BG_COLOR)
    api_frame.pack(fill='x', pady=(0, 15))
    
    api_label = Label(api_frame,
                     text="OpenAI API Key:",
                     font=('Arial', 11),
                     bg=BG_COLOR,
                     fg=TEXT_COLOR)
    api_label.pack(anchor='w')
    
    key_entry = ttk.Entry(api_frame,
                         style='Modern.TEntry',
                         width=50)
    key_entry.pack(fill='x', pady=(5, 0))
    key_entry.insert(0, current_key)
    
    # API Base URL section
    base_frame = Frame(settings_container, bg=BG_COLOR)
    base_frame.pack(fill='x', pady=(0, 15))
    
    base_label = Label(base_frame,
                      text="API Base URL:",
                      font=('Arial', 11),
                      bg=BG_COLOR,
                      fg=TEXT_COLOR)
    base_label.pack(anchor='w')
    
    base_value = Label(base_frame,
                      text=OPENAI_API_BASE,
                      font=('Arial', 10),
                      bg='white',
                      fg=TEXT_COLOR,
                      relief='flat',
                      padx=10,
                      pady=5)
    base_value.pack(fill='x', pady=(5, 0))
    
    # Password section
    pass_frame = Frame(settings_container, bg=BG_COLOR)
    pass_frame.pack(fill='x', pady=(0, 15))
    
    pass_label = Label(pass_frame,
                      text="New Password:",
                      font=('Arial', 11),
                      bg=BG_COLOR,
                      fg=TEXT_COLOR)
    pass_label.pack(anchor='w')
    
    new_entry = ttk.Entry(pass_frame,
                         style='Modern.TEntry',
                         show='*',
                         width=50)
    new_entry.pack(fill='x', pady=(5, 0))
    
    conf_label = Label(pass_frame,
                      text="Confirm Password:",
                      font=('Arial', 11),
                      bg=BG_COLOR,
                      fg=TEXT_COLOR)
    conf_label.pack(anchor='w', pady=(10, 0))
    
    conf_entry = ttk.Entry(pass_frame,
                          style='Modern.TEntry',
                          show='*',
                          width=50)
    conf_entry.pack(fill='x', pady=(5, 0))
    
    # Save button
    save_btn = ttk.Button(settings_container,
                         text="💾 Save Settings",
                         style='Modern.TButton',
                         command=save)
    save_btn.pack(pady=(20, 0))
    
    # Add hover effects
    def on_enter(e):
        e.widget.configure(style='Modern.TButton')
    
    def on_leave(e):
        e.widget.configure(style='Modern.TButton')
    
    save_btn.bind('<Enter>', on_enter)
    save_btn.bind('<Leave>', on_leave)

def open_settings():
    verify_password()

def audio_callback(indata, frames_count, time_info, status):
    if status:
        log(f'Audio status: {status}', level='WARNING')
    frames.append(indata.copy())

# Add a status bar at the bottom of the GUI
status_label = Label(root, text="Ready", anchor="w", font=("Arial", 10), bg="#f0f0f0", relief="sunken")
status_label.pack(side='bottom', fill='x', padx=5, pady=5)

def update_status(message):
    """Update the status bar with the given message."""
    status_label.config(text=message)

def start_recording():
    global recording, frames, tts_engine
    try:
        # Stop any ongoing speech
        if is_speaking and tts_engine is not None:
            tts_engine.stop()
        
        # Clear previous frames
        frames.clear()
        recording.set()
        
        # Update UI
        start_btn.config(state='disabled')
        stop_btn.config(state='normal')
        update_status("Recording started... Click 'Stop' when done.")
        
        # Change only status label background color to light red
        status_label.configure(bg='#FFE4E1')  # Light red color
        
        # Start recording in a separate thread
        threading.Thread(target=run_stream, daemon=True).start()
    except Exception as e:
        print(f"Error starting recording: {e}")
        update_status('Error starting recording')

def stop_recording():
    global recording
    try:
        recording.clear()
        start_btn.config(state='normal')
        stop_btn.config(state='disabled')
        update_status("Recording stopped.")
        
        # Reset status label background color to default
        status_label.configure(bg='#f0f0f0')
        
        # Process the recorded audio
        process_audio()
    except Exception as e:
        print(f"Error stopping recording: {e}")
        update_status('Error stopping recording')

def run_stream():
    global stream
    try:
        stream = sd.InputStream(samplerate=16000, channels=1, callback=audio_callback)
        with stream:
            while recording.is_set():
                sd.sleep(100)
    except Exception as e:
        print(f'Error starting recording: {e}')
    finally:
        stream = None
        stop_btn.config(state='disabled')
        start_btn.config(state='normal')

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

def correct_grammar_gpt(text):
    try:
        if not openai.api_key:
            return "Error: OpenAI API key not set. Please configure it in Settings."
            
        client = openai.OpenAI(
            api_key=openai.api_key,
            base_url=OPENAI_API_BASE
        )
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an English grammar assistant."},
                {"role": "user", "content": text}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"OpenAI API Error: {str(e)}")
        return f'Error: {str(e)}'

def initialize_tts():
    global tts_engine
    try:
        tts_engine = pyttsx3.init()
        return True
    except Exception as e:
        print(f"Error initializing TTS: {e}")
        return False

def play_voice(text):
    global is_speaking, tts_engine
    try:
        is_speaking = True
        if tts_engine is None:
            if not initialize_tts():
                return
        
        def speak():
            try:
                tts_engine.say(text)
                tts_engine.runAndWait()
            except Exception as e:
                print(f"Error in TTS thread: {e}")
            finally:
                global is_speaking
                is_speaking = False
        
        # Run TTS in a separate thread
        tts_thread = threading.Thread(target=speak, daemon=True)
        tts_thread.start()
    except Exception as e:
        print(f"Error starting TTS thread: {e}")
        is_speaking = False

def cleanup_tts():
    global tts_engine, is_speaking
    try:
        if tts_engine is not None:
            if is_speaking:
                tts_engine.stop()
            tts_engine = None
        is_speaking = False
    except Exception as e:
        print(f"Error cleaning up TTS: {e}")

def initialize_grammar_model():
    """Initialize the grammar correction model."""
    global grammar_model, tokenizer
    try:
        if use_online_mode.get():
            # Load a pre-trained grammar correction model
            model_name = "deep-learning-analytics/writing-assistant"
            tokenizer = AutoTokenizer.from_pretrained(model_name)
            grammar_model = AutoModelForSeq2SeqLM.from_pretrained(model_name)
            return True
        return False
    except Exception as e:
        print(f"Error initializing grammar model: {e}")
        return False

class GrammarCorrector:
    """Advanced grammar correction system with dual-mode support."""
    
    def __init__(self):
        # Basic grammar rules for offline mode
        self.basic_rules = {
            # Subject-verb agreement
            r'\b(I|you|we|they)\s+(is|was)\b': lambda m: f"{m.group(1)} {'am' if m.group(1)=='I' else 'are' if m.group(2)=='is' else 'were'}",
            r'\b(he|she|it)\s+are\b': lambda m: f"{m.group(1)} is",
            r'\b(he|she|it)\s+were\b': lambda m: f"{m.group(1)} was",
            
            # Articles
            r'\b(a)\s+([aeiouAEIOU][a-zA-Z]*)\b': lambda m: f"an {m.group(2)}",
            r'\b(an)\s+([^aeiouAEIOU][a-zA-Z]*)\b': lambda m: f"a {m.group(2)}",
            
            # Common mistakes
            r'\b(could|would|should|must|will|can)\s+of\b': lambda m: f"{m.group(1)} have",
            r'\bin\s+regards\s+to\b': "regarding",
            r'\bless\s+([a-zA-Z]+s)\b': lambda m: f"fewer {m.group(1)}",
            r'\bmore\s+better\b': "better",
            r'\bmost\s+biggest\b': "biggest",
            r'\bvery\s+unique\b': "unique",
            
            # Punctuation
            r'\s+([.,!?:;])': r'\1',
            r'([.,!?:;])([^\s])': r'\1 \2',
            r'\s+': ' ',
        }
        
        # Advanced rules for offline mode
        self.advanced_rules = {
            # Complex sentence structure
            r'\b(however|therefore|moreover|furthermore|nevertheless)\s+([a-z])': lambda m: f"{m.group(1)}, {m.group(2).upper()}",
            r'\b(in\s+addition|for\s+example|in\s+fact|in\s+other\s+words)\s+([a-z])': lambda m: f"{m.group(1)}, {m.group(2).upper()}",
            
            # Word choice improvements
            r'\b(affect|effect)\b': self.fix_affect_effect,
            r'\b(accept|except)\b': self.fix_accept_except,
            r'\b(advice|advise)\b': self.fix_advice_advise,
            
            # Tense consistency
            r'\b(yesterday|last week|last month|ago)\s+([a-zA-Z]*(?:s|es|ies))\b': lambda m: f"{m.group(1)} {self.to_past_tense(m.group(2))}",
            r'\b(tomorrow|next week|next month)\s+([a-zA-Z]*ed)\b': lambda m: f"{m.group(1)} {self.to_future_tense(m.group(2))}",
        }

    def correct_text(self, text):
        """Apply grammar corrections based on internet availability."""
        if check_internet() and openai.api_key:
            return self.correct_text_online(text)
        else:
            return self.correct_text_offline(text)

    def correct_text_online(self, text):
        """Use OpenAI for grammar correction."""
        try:
            print(f"Using API Key: {openai.api_key[:10]}...")
            print(f"Using API Base: {OPENAI_API_BASE}")
            
            # Configure OpenAI client
            client = openai.OpenAI(
                api_key=openai.api_key,
                base_url=OPENAI_API_BASE
            )
            
            # Make the API call
            response = client.chat.completions.create(
                model="openai/gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are an English grammar expert. Correct any grammar, spelling, or punctuation errors in the text while maintaining the original meaning. Keep the response concise and clear."},
                    {"role": "user", "content": text}
                ],
                temperature=0.3,
                max_tokens=500
            )
            
            print("API Response received successfully")
            return response.choices[0].message.content.strip()
        except Exception as e:
            print(f"Error in online correction: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            return self.correct_text_offline(text)

    def correct_text_offline(self, text):
        """Apply offline grammar corrections."""
        # Split into sentences
        sentences = re.split(r'([.!?]+\s+)', text)
        corrected_sentences = []
        
        for i in range(0, len(sentences), 2):
            sentence = sentences[i]
            if i + 1 < len(sentences):
                sentence += sentences[i + 1]
            
            # Apply basic rules
            for pattern, replacement in self.basic_rules.items():
                sentence = re.sub(pattern, replacement, sentence, flags=re.IGNORECASE)
            
            # Apply advanced rules
            for pattern, replacement in self.advanced_rules.items():
                sentence = re.sub(pattern, replacement, sentence, flags=re.IGNORECASE)
            
            # Capitalize first letter
            if sentence:
                sentence = sentence[0].upper() + sentence[1:]
            
            corrected_sentences.append(sentence)
        
        # Join and clean up
        text = ''.join(corrected_sentences)
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\s*([.,!?:;])', r'\1', text)
        text = re.sub(r'([.,!?:;])(?=[^\s])', r'\1 ', text)
        return text.strip()

    def to_past_tense(self, word):
        """Convert a word to past tense."""
        irregular = {
            'go': 'went', 'have': 'had', 'do': 'did', 'say': 'said',
            'make': 'made', 'take': 'took', 'come': 'came', 'see': 'saw',
            'know': 'knew', 'get': 'got'
        }
        
        if word.lower() in irregular:
            return irregular[word.lower()]
        elif word.endswith('e'):
            return word + 'd'
        elif word.endswith('y'):
            return word[:-1] + 'ied'
        else:
            return word + 'ed'

    def to_future_tense(self, word):
        """Convert a word to future tense."""
        if word.endswith('ed'):
            if word.endswith('ied'):
                return word[:-3] + 'y'
            return word[:-2]
        return word

    def fix_affect_effect(self, match):
        """Context-aware affect/effect correction."""
        word = match.group(1)
        prev_words = match.string[:match.start()].split()[-3:]
        
        if any(w in prev_words for w in ['the', 'an', 'any', 'no']):
            return 'effect'
        if any(w in prev_words for w in ['to', 'will', 'can', 'may']):
            return 'affect'
        return word

    def fix_accept_except(self, match):
        """Context-aware accept/except correction."""
        word = match.group(1)
        prev_words = match.string[:match.start()].split()[-3:]
        
        if any(w in prev_words for w in ['to', 'will', 'can', 'may']):
            return 'accept'
        if any(w in prev_words for w in ['for', 'but']):
            return 'except'
        return word

    def fix_advice_advise(self, match):
        """Context-aware advice/advise correction."""
        word = match.group(1)
        prev_words = match.string[:match.start()].split()[-3:]
        
        if any(w in prev_words for w in ['the', 'some', 'any', 'good']):
            return 'advice'
        if any(w in prev_words for w in ['to', 'will', 'can', 'may']):
            return 'advise'
        return word

# Create a global instance of the grammar corrector
grammar_corrector = GrammarCorrector()

def correct_grammar_offline(text):
    """Advanced grammar checking using the GrammarCorrector class."""
    return grammar_corrector.correct_text(text)

# Add this before the menu creation
def check_api_status():
    """Check and display API connection status."""
    status_window = Toplevel(root)
    status_window.title('API Connection Status')
    status_window.geometry('400x200')
    
    # Center the window
    status_window.update_idletasks()
    width = status_window.winfo_width()
    height = status_window.winfo_height()
    x = (status_window.winfo_screenwidth() // 2) - (width // 2)
    y = (status_window.winfo_screenheight() // 2) - (height // 2)
    status_window.geometry(f'{width}x{height}+{x}+{y}')
    
    # Create status frame
    status_frame = Frame(status_window, padx=20, pady=20)
    status_frame.pack(expand=True, fill='both')
    
    # Check internet connection
    internet_status = "Connected" if check_internet() else "Not Connected"
    internet_label = Label(status_frame, text=f"Internet: {internet_status}")
    internet_label.pack(pady=5)
    
    # Check API key
    api_key_status = "Set" if openai.api_key else "Not Set"
    api_key_label = Label(status_frame, text=f"API Key: {api_key_status}")
    api_key_label.pack(pady=5)
    
    # Status label
    status_label = Label(status_frame, text="Testing connection...", fg="blue")
    status_label.pack(pady=10)
    
    # Test button
    test_button = Button(status_frame, text="Test Connection", command=lambda: test_connection(status_label))
    test_button.pack(pady=10)
    
    # Close button
    close_button = Button(status_frame, text="Close", command=status_window.destroy)
    close_button.pack(pady=10)
    
    # Test connection immediately
    test_connection(status_label)

def test_connection(status_label):
    """Test the API connection and update the status label."""
    try:
        if not openai.api_key:
            status_label.config(text="API Key not set. Please configure it in Settings.", fg="red")
            return
            
        client = openai.OpenAI(
            api_key=openai.api_key,
            base_url=OPENAI_API_BASE
        )
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an English grammar expert."},
                {"role": "user", "content": "Test"}
            ],
            max_tokens=5
        )
        status_label.config(text="API Connection: Successful", fg="green")
    except Exception as e:
        status_label.config(text=f"API Connection: Failed - {str(e)}", fg="red")

def check_microphone_status():
    """Check if microphone is available and working."""
    try:
        devices = sd.query_devices()
        input_devices = [d for d in devices if d['max_input_channels'] > 0]
        if not input_devices:
            return False, "No microphone detected"
        
        # Try to open a test stream
        test_stream = sd.InputStream(samplerate=16000, channels=1)
        test_stream.close()
        return True, None
    except Exception as e:
        return False, str(e)

# Add this after the imports and before the menu creation
def show_about_window():
    about_window = Toplevel(root)
    about_window.title('About Us')
    about_window.geometry('400x500')
    about_window.resizable(False, False)
    about_window.configure(bg=BG_COLOR)
    
    # Center the window
    about_window.update_idletasks()
    width = about_window.winfo_width()
    height = about_window.winfo_height()
    x = (about_window.winfo_screenwidth() // 2) - (width // 2)
    y = (about_window.winfo_screenheight() // 2) - (height // 2)
    about_window.geometry(f'{width}x{height}+{x}+{y}')
    
    # Create main frame
    main_frame = Frame(about_window, bg=BG_COLOR, padx=30, pady=30)
    main_frame.pack(fill='both', expand=True)
    
    # Add title
    title_label = Label(main_frame,
                       text="ℹ️ About",
                       font=('Arial', 16, 'bold'),
                       bg=BG_COLOR,
                       fg=PRIMARY_COLOR)
    title_label.pack(pady=(0, 20))
    
    # Load and display logo
    try:
        logo_img = Image.open(app_icon_path)
        logo_img = logo_img.resize((120, 120), Image.LANCZOS)
        logo_photo = ImageTk.PhotoImage(logo_img)
        logo_label = Label(main_frame, image=logo_photo, bg=BG_COLOR)
        logo_label.image = logo_photo  # Keep a reference
        logo_label.pack(pady=(0, 20))
    except Exception as e:
        print(f"Error loading logo: {e}")
    
    # Developer info
    dev_frame = Frame(main_frame, bg=BG_COLOR)
    dev_frame.pack(fill='x', pady=(0, 20))
    
    name_label = Label(dev_frame,
                      text="Md. Yamin Hossain",
                      font=('Arial', 14, 'bold'),
                      bg=BG_COLOR,
                      fg=TEXT_COLOR)
    name_label.pack(pady=(0, 5))
    
    # Website section
    website_frame = Frame(main_frame, bg=BG_COLOR)
    website_frame.pack(fill='x', pady=(0, 20))
    
    website_label = Label(website_frame,
                         text="🌐 Website",
                         font=('Arial', 12),
                         bg=BG_COLOR,
                         fg=TEXT_COLOR)
    website_label.pack(pady=(0, 5))
    
    def open_website():
        import webbrowser
        webbrowser.open('https://needyamin.github.io')
    
    website_link = Label(website_frame,
                        text="https://needyamin.github.io",
                        font=('Arial', 11),
                        fg=PRIMARY_COLOR,
                        bg=BG_COLOR,
                        cursor='hand2')
    website_link.pack()
    website_link.bind('<Button-1>', lambda e: open_website())
    
    # Version info
    version_frame = Frame(main_frame, bg=BG_COLOR)
    version_frame.pack(fill='x', pady=(0, 20))
    
    version_label = Label(version_frame,
                         text="📦 Version 1.0",
                         font=('Arial', 11),
                         bg=BG_COLOR,
                         fg=TEXT_COLOR)
    version_label.pack()
    
    # Close button
    close_btn = ttk.Button(main_frame,
                          text="✕ Close",
                          style='Modern.TButton',
                          command=about_window.destroy)
    close_btn.pack(pady=(20, 0))
    
    # Add hover effects
    def on_enter(e):
        e.widget.configure(style='Modern.TButton')
    
    def on_leave(e):
        e.widget.configure(style='Modern.TButton')
    
    close_btn.bind('<Enter>', on_enter)
    close_btn.bind('<Leave>', on_leave)

# Add this function after show_about_window
def show_setup_instructions():
    setup_window = Toplevel(root)
    setup_window.title('How to Setup')
    setup_window.geometry('600x500')
    setup_window.resizable(False, False)
    setup_window.configure(bg=BG_COLOR)
    
    # Center the window
    setup_window.update_idletasks()
    width = setup_window.winfo_width()
    height = setup_window.winfo_height()
    x = (setup_window.winfo_screenwidth() // 2) - (width // 2)
    y = (setup_window.winfo_screenheight() // 2) - (height // 2)
    setup_window.geometry(f'{width}x{height}+{x}+{y}')
    
    # Create main frame
    main_frame = Frame(setup_window, bg=BG_COLOR, padx=30, pady=30)
    main_frame.pack(fill='both', expand=True)
    
    # Add title
    title_label = Label(main_frame,
                       text="📝 Setup Instructions",
                       font=('Arial', 16, 'bold'),
                       bg=BG_COLOR,
                       fg=PRIMARY_COLOR)
    title_label.pack(pady=(0, 20))
    
    # Create a frame for the instructions
    instructions_frame = Frame(main_frame, bg=BG_COLOR)
    instructions_frame.pack(fill='both', expand=True)
    
    # Setup instructions with proper formatting
    instructions = [
        "1. Go to https://openrouter.ai/ and create your own account.",
        "2. Then go to https://openrouter.ai/settings/keys and create your own key.",
        "3. Copy your key.",
        "4. Back on the software, go to File > Settings",
        "5. Type password: admin",
        "6. Replace the old API with your new API",
        "7. Enter your new password and save",
        "8. Done! You are ready to fly!"
    ]
    
    # Add instructions with proper formatting
    for i, instruction in enumerate(instructions):
        step_frame = Frame(instructions_frame, bg=BG_COLOR)
        step_frame.pack(fill='x', pady=5)
        
        # Step number
        step_label = Label(step_frame,
                          text=f"{i+1}.",
                          font=('Arial', 11, 'bold'),
                          bg=BG_COLOR,
                          fg=PRIMARY_COLOR,
                          width=3)
        step_label.pack(side='left', padx=(0, 5))
        
        # Instruction text
        instruction_label = Label(step_frame,
                                text=instruction,
                                font=('Arial', 11),
                                bg=BG_COLOR,
                                fg=TEXT_COLOR,
                                wraplength=500,
                                justify='left',
                                anchor='w')
        instruction_label.pack(side='left', fill='x', expand=True)
    
    # Add clickable links
    def open_openrouter():
        import webbrowser
        webbrowser.open('https://openrouter.ai/')
    
    def open_keys():
        import webbrowser
        webbrowser.open('https://openrouter.ai/settings/keys')
    
    # Create links frame
    links_frame = Frame(main_frame, bg=BG_COLOR)
    links_frame.pack(fill='x', pady=(20, 0))
    
    # OpenRouter link
    link1_frame = Frame(links_frame, bg=BG_COLOR)
    link1_frame.pack(fill='x', pady=5)
    
    link1_label = Label(link1_frame,
                       text="🔗 OpenRouter Website:",
                       font=('Arial', 11),
                       bg=BG_COLOR,
                       fg=TEXT_COLOR)
    link1_label.pack(side='left', padx=(0, 5))
    
    link1 = Label(link1_frame,
                 text="https://openrouter.ai/",
                 font=('Arial', 11),
                 fg=PRIMARY_COLOR,
                 bg=BG_COLOR,
                 cursor='hand2')
    link1.pack(side='left')
    link1.bind('<Button-1>', lambda e: open_openrouter())
    
    # Keys link
    link2_frame = Frame(links_frame, bg=BG_COLOR)
    link2_frame.pack(fill='x', pady=5)
    
    link2_label = Label(link2_frame,
                       text="🔑 API Keys Page:",
                       font=('Arial', 11),
                       bg=BG_COLOR,
                       fg=TEXT_COLOR)
    link2_label.pack(side='left', padx=(0, 5))
    
    link2 = Label(link2_frame,
                 text="https://openrouter.ai/settings/keys",
                 font=('Arial', 11),
                 fg=PRIMARY_COLOR,
                 bg=BG_COLOR,
                 cursor='hand2')
    link2.pack(side='left')
    link2.bind('<Button-1>', lambda e: open_keys())
    
    # Close button
    close_btn = ttk.Button(main_frame,
                          text="✕ Close",
                          style='Modern.TButton',
                          command=setup_window.destroy)
    close_btn.pack(pady=(20, 0))
    
    # Add hover effects
    def on_enter(e):
        e.widget.configure(style='Modern.TButton')
    
    def on_leave(e):
        e.widget.configure(style='Modern.TButton')
    
    close_btn.bind('<Enter>', on_enter)
    close_btn.bind('<Leave>', on_leave)

# Then create the menu
menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label='Settings', command=open_settings)
file_menu.add_command(label='Clear Logs', command=lambda: log('Clear Logs'))
file_menu.add_separator()
file_menu.add_command(label='Exit', command=root.quit)
menu_bar.add_cascade(label='File', menu=file_menu)

# Create Help menu
help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label='API Status', command=check_api_status)
help_menu.add_separator()
help_menu.add_command(label='How to Setup?', command=show_setup_instructions)
help_menu.add_command(label='About', command=show_about_window)
menu_bar.add_cascade(label='Help', menu=help_menu)

# Configure the menu bar
root.config(menu=menu_bar)

# Create main container
main_container = Frame(root, bg=BG_COLOR)
main_container.pack(fill='both', expand=True, padx=20, pady=20)

# Create header frame
header_frame = Frame(main_container, bg=BG_COLOR)
header_frame.pack(fill='x', pady=(0, 20))

# Add title label
title_label = Label(header_frame,
                   text="🎧 Speech Assistant",
                   font=("Arial", 16, "bold"),
                   bg=BG_COLOR,
                   fg=TEXT_COLOR)
title_label.pack(side='left')

# Add status indicator
status_indicator = Canvas(header_frame, width=20, height=20, bg=BG_COLOR, highlightthickness=0)
status_indicator.pack(side='right', padx=(0, 10))
status_circle = status_indicator.create_oval(2, 2, 18, 18, fill=SUCCESS_COLOR)

# Create chat area with modern styling
chat_frame = Frame(main_container, bg=BG_COLOR)
chat_frame.pack(fill='both', expand=True)

# Create scrollable text area with modern styling
output = scrolledtext.ScrolledText(chat_frame,
                                 wrap=tk.WORD,
                                 state='disabled',
                                 font=('Arial', 11),
                                 bg='white',
                                 fg=TEXT_COLOR,
                                 padx=15,
                                 pady=15,
                                 relief='flat',
                                 highlightthickness=0)
output.pack(fill='both', expand=True)

# Create control panel with modern styling
control_panel = Frame(main_container, bg=BG_COLOR)
control_panel.pack(fill='x', pady=(20, 0))

# Create button container
btn_container = Frame(control_panel, bg=BG_COLOR)
btn_container.pack(side='left', fill='x', expand=True)

# Create modern buttons
start_btn = ttk.Button(btn_container,
                      text='🎤 Start Recording',
                      style='Modern.TButton',
                      command=start_recording)
start_btn.pack(side='left', padx=5)

stop_btn = ttk.Button(btn_container,
                     text='⏹️ Stop Recording',
                     style='Modern.TButton',
                     state='disabled',
                     command=stop_recording)
stop_btn.pack(side='left', padx=5)

clear_btn = ttk.Button(btn_container,
                      text='🗑️ Clear Logs',
                      style='Modern.TButton',
                      command=lambda: log('Clear Logs'))
clear_btn.pack(side='left', padx=5)

settings_btn = ttk.Button(btn_container,
                        text='⚙️ Settings',
                        style='Modern.TButton',
                        command=open_settings)
settings_btn.pack(side='left', padx=5)

# Get available microphone devices
def get_mic_devices():
    try:
        devices = sd.query_devices()
        mic_list = [d['name'] for d in devices if d['max_input_channels'] > 0]
        return mic_list if mic_list else ['Default Microphone']
    except Exception as e:
        print(f"Error getting microphone devices: {e}")
        return ['Default Microphone']

# Initialize microphone devices
mic_devices = get_mic_devices()

# Create microphone selector with modern styling
mic_frame = Frame(control_panel, bg=BG_COLOR)
mic_frame.pack(side='right', fill='x', padx=5)

mic_label = Label(mic_frame,
                 text="🎤 Microphone:",
                 font=('Arial', 10),
                 bg=BG_COLOR,
                 fg=TEXT_COLOR)
mic_label.pack(side='left', padx=(0, 5))

mic_dropdown = ttk.Combobox(mic_frame,
                          values=mic_devices,
                          style='Modern.TCombobox',
                          state='readonly',
                          width=30)
mic_dropdown.pack(side='left')
if mic_devices:
    mic_dropdown.set(mic_devices[0])

# Create status bar with modern styling
status_bar = Frame(main_container, bg=BG_COLOR, height=30)
status_bar.pack(fill='x', pady=(10, 0))

status_label = Label(status_bar,
                    text="Ready",
                    font=('Arial', 9),
                    bg=BG_COLOR,
                    fg=TEXT_COLOR)
status_label.pack(side='left', padx=5)

# Configure grid weights
main_container.grid_rowconfigure(1, weight=1)
main_container.grid_columnconfigure(0, weight=1)

# Tray integration
def center_window(window):
    """Center the window on the screen."""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def setup_tray():
    global tray_icon, tray_icon_initialized
    try:
        # Only create tray icon if it doesn't exist
        if not tray_icon_initialized:
            # Create a default icon if the file doesn't exist
            if not os.path.exists(app_icon_path):
                img = Image.new('RGB', (64, 64), color='blue')
            else:
                img = Image.open(app_icon_path)
            
            # Create the tray icon
            menu = (
                item('Restore', restore),
                item('Quit', quit_app)
            )
            
            # Stop any existing tray icon
            if tray_icon is not None:
                try:
                    tray_icon.stop()
                except:
                    pass
            
            tray_icon = Icon('Speech Assistant', img, 'Speech Assistant', menu)
            tray_icon_initialized = True
            
            # Start the tray icon in a separate thread
            threading.Thread(target=tray_icon.run, daemon=True).start()
            return True
        return True
    except Exception as e:
        print(f"Error setting up tray icon: {e}")
        return False

def on_closing():
    try:
        # Clean up TTS first
        cleanup_tts()
        
        # Hide the window
        root.withdraw()
        
        # Ensure tray icon is set up
        if not tray_icon_initialized:
            setup_tray()
        
        # Show notification
        if tray_icon_initialized and tray_icon:
            try:
                tray_icon.notify('Speech Assistant is running in the background', 'Click the tray icon to restore')
            except:
                pass
    except Exception as e:
        print(f"Error minimizing to tray: {e}")

def on_minimize(event):
    try:
        # Allow minimizing if user explicitly clicks minimize button
        if event.widget == root:
            # Hide the window
            root.withdraw()
            
            # Ensure tray icon is set up
            if not tray_icon_initialized:
                setup_tray()
            
            # Show notification
            if tray_icon_initialized and tray_icon:
                try:
                    tray_icon.notify('Speech Assistant is running in the background', 'Click the tray icon to restore')
                except:
                    pass
    except Exception as e:
        print(f"Error minimizing window: {e}")

def restore(icon, item):
    try:
        # Show the window
        root.deiconify()
        root.state('normal')
        root.lift()
        root.focus_force()
    except Exception as e:
        print(f"Error restoring window: {e}")

def quit_app(icon, item):
    try:
        # Clean up TTS first
        cleanup_tts()
        
        # Stop the tray icon
        if tray_icon_initialized and tray_icon:
            try:
                tray_icon.stop()
            except:
                pass
        
        # Quit the application
        root.quit()
        root.destroy()
    except Exception as e:
        print(f"Error quitting application: {e}")
        root.quit()
        root.destroy()

# Initialize tray icon at startup
setup_tray()

# Update the window protocol
root.protocol('WM_DELETE_WINDOW', on_closing)
root.bind('<Unmap>', on_minimize)

# Center the main window
center_window(root)

# Initialize loading animation
if not load_animation():
    log('Warning: Loading animation not available.', level='WARNING')

# Check internet connection
if not check_internet():
    root.after(1000, show_no_internet_dialog)

# Paths for shortcuts
startup_folder = winshell.startup()
startup_shortcut_path = os.path.join(startup_folder, "Speech Assistant.lnk")
start_menu_shortcut_path = os.path.join(winshell.start_menu(), "Speech Assistant.lnk")

def create_startup_shortcut():
    """Create a shortcut for the app in the Windows Startup folder."""
    try:
        if not os.path.exists(startup_shortcut_path):
            shell = Dispatch("WScript.Shell", pythoncom.CoInitialize())
            shortcut = shell.CreateShortcut(startup_shortcut_path)
            shortcut.TargetPath = os.path.abspath(sys.argv[0])  # Use absolute path of the script
            shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[0]))
            shortcut.IconLocation = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "needyamin.ico")
            shortcut.Save()
            print(f"Startup shortcut created at: {startup_shortcut_path}")
    except Exception as e:
        print(f"Error creating startup shortcut: {e}")

def create_start_menu_shortcut():
    """Create a shortcut for the app in the Windows Start Menu."""
    try:
        if not os.path.exists(start_menu_shortcut_path):
            shell = Dispatch("WScript.Shell", pythoncom.CoInitialize())
            shortcut = shell.CreateShortcut(start_menu_shortcut_path)
            shortcut.TargetPath = os.path.abspath(sys.argv[0])  # Use absolute path of the script
            shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[0]))
            shortcut.IconLocation = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "needyamin.ico")
            shortcut.Save()
            print(f"Start Menu shortcut created at: {start_menu_shortcut_path}")
    except Exception as e:
        print(f"Error creating Start Menu shortcut: {e}")

def remove_startup_shortcut():
    """Remove the startup shortcut."""
    try:
        if os.path.exists(startup_shortcut_path):
            os.remove(startup_shortcut_path)
            print(f"Startup shortcut removed: {startup_shortcut_path}")
    except Exception as e:
        print(f"Error removing startup shortcut: {e}")

def is_startup_enabled():
    """Check if the startup shortcut exists."""
    return os.path.exists(startup_shortcut_path)

def update_startup_menu():
    """Update the Help menu to show the current startup status with tick icons."""
    help_menu.delete(0, "end")  # Clear existing menu items
    
    # Add API Status
    help_menu.add_command(label='API Status', command=check_api_status)
    help_menu.add_separator()
    
    # Add How to Setup
    help_menu.add_command(label='How to Setup?', command=show_setup_instructions)
    help_menu.add_separator()
    
    # Add startup options
    if is_startup_enabled():
        help_menu.add_command(label="✓ Start on Startup", command=lambda: toggle_startup(True))
        help_menu.add_command(label="Stop on Startup", command=lambda: toggle_startup(False))
    else:
        help_menu.add_command(label="Start on Startup", command=lambda: toggle_startup(True))
        help_menu.add_command(label="✓ Stop on Startup", command=lambda: toggle_startup(False))
    
    help_menu.add_separator()
    help_menu.add_command(label="About", command=show_about_window)

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

# Initialize TTS engine at startup
initialize_tts()

# Start the application
if __name__ == '__main__':
    try:
        root.mainloop()
    except Exception as e:
        print(f"Error in main loop: {e}")
    finally:
        cleanup_tts()
        if tray_icon_initialized and tray_icon:
            tray_icon.stop()

def create_desktop_shortcut():
    """Create a shortcut for the app on the Windows Desktop."""
    try:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        desktop_shortcut_path = os.path.join(desktop_path, "Speech Assistant.lnk")
        
        if not os.path.exists(desktop_shortcut_path):
            shell = Dispatch("WScript.Shell", pythoncom.CoInitialize())
            shortcut = shell.CreateShortcut(desktop_shortcut_path)
            shortcut.TargetPath = os.path.abspath(sys.argv[0])  # Use absolute path of the script
            shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[0]))
            shortcut.IconLocation = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "needyamin.ico")
            shortcut.Save()
            print(f"Desktop shortcut created at: {desktop_shortcut_path}")
    except Exception as e:
        print(f"Error creating desktop shortcut: {e}")

# Create all shortcuts
create_startup_shortcut()
create_start_menu_shortcut()
create_desktop_shortcut()