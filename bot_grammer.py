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
root.geometry('700x650')

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
    
    bubble_color = '#E8E8E8' if is_bot else '#DCF8C6'
    bubble_frame, content_frame = create_rounded_frame(frame, bubble_color)
    bubble_frame.pack(pady=5, padx=10, anchor='w' if is_bot else 'e', fill='x')
    
    # Time and copy button frame
    time_frame = Frame(content_frame, bg=bubble_color)
    time_frame.pack(fill='x', padx=5, pady=(5,0))
    
    # Time label
    time_label = Label(time_frame, 
                      text=time.strftime('%d %B %Y: %I:%M %p').lstrip('0'),
                      font=('Consolas', 8),
                      fg='gray',
                      bg=bubble_color)
    time_label.pack(side='left', padx=(0,10))
    
    # Copy button
    copy_button = Button(time_frame,
                        text='📋 Copy',
                        font=('Consolas', 8),
                        bg='#4CAF50',
                        fg='white',
                        relief='flat',
                        cursor='hand2',
                        width=8)
    copy_button.pack(side='left')
    
    # Message content frame
    msg_content = Frame(content_frame, bg=bubble_color)
    msg_content.pack(fill='x', padx=5, pady=5)
    
    # Icon (if provided)
    if icon_image:
        icon_label = Label(msg_content, image=icon_image, bg=bubble_color)
        icon_label.pack(side='left', padx=(0,5))
    
    # Message text frame
    text_frame = Frame(msg_content, bg=bubble_color)
    text_frame.pack(side='left', fill='x', expand=True)
    
    # Create a text widget for the message
    msg_text = Text(text_frame,
                   wrap='word',
                   width=60,
                   height=1,
                   font=('Consolas', 11),
                   bg=bubble_color,
                   relief='flat',
                   padx=5,
                   pady=5)
    msg_text.pack(side='left', fill='x', expand=True)
    msg_text.insert('1.0', message)
    msg_text.configure(state='disabled')
    
    # Adjust height based on content
    def adjust_height():
        # Count the number of lines in the text
        line_count = msg_text.count('1.0', 'end', 'lines')[0]
        # Set the height to match the content
        msg_text.configure(height=line_count)
        # Update the canvas height
        canvas = bubble_frame.winfo_children()[0]
        # Calculate required height: line_count * line_height + padding
        required_height = line_count * 20 + 60  # 20 pixels per line + padding
        canvas.configure(height=required_height)
        # Update the frame to accommodate the new height
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
                               text='✓ Copied!',
                               font=('Consolas', 8),
                               fg='green',
                               bg=bubble_color)
            copied_label.pack(side='left', padx=(5,0))
            
            # Animate the copied label
            def fade_out():
                copied_label.configure(fg='#00FF00')
                time_frame.after(500, lambda: copied_label.configure(fg='green'))
                time_frame.after(1000, copied_label.destroy)
            
            fade_out()
        except Exception as e:
            print(f"Error copying text: {e}")
    
    # Bind the copy function to the button
    copy_button.configure(command=copy_text)
    
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
    
    # Get current settings
    cur.execute('SELECT api_key, password FROM config WHERE id=1')
    current_settings = cur.fetchone()
    current_key = current_settings[0] if current_settings else ''
    
    Label(sett, text='OpenAI API Key:').grid(row=0, column=0, padx=5, pady=5)
    key_entry = Entry(sett, width=50)
    key_entry.grid(row=0, column=1, padx=5, pady=5)
    key_entry.insert(0, current_key)
    
    Label(sett, text='API Base URL:').grid(row=1, column=0, padx=5, pady=5)
    base_label = Label(sett, text=OPENAI_API_BASE, width=50)
    base_label.grid(row=1, column=1, padx=5, pady=5)
    
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
        tts_engine.say(text)
        tts_engine.runAndWait()
    except Exception as e:
        print(f"Error playing voice: {e}")
    finally:
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

# Then create the menu
menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label='Settings', command=open_settings)
file_menu.add_command(label='Clear Logs', command=lambda: log('Clear Logs'))
file_menu.add_separator()
file_menu.add_command(label='Exit', command=root.quit)
menu_bar.add_cascade(label='File', menu=file_menu)

help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label='API Status', command=check_api_status)
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
mic_dropdown.grid(row=0, column=4, padx=5, sticky='ew')
mic_dropdown.grid_remove()  # Hide by default

# Check microphone status and show dropdown only if needed
mic_status, mic_error = check_microphone_status()
if not mic_status:
    mic_dropdown.grid()  # Show dropdown if there's an error
    log(f"Microphone error: {mic_error}", level='WARNING')
elif mic_devices:
    mic_dropdown.set(mic_devices[0])

# Configure grid weights
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

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
        if tray_icon_initialized:
            tray_icon.notify('Speech Assistant is running in the background', 'Click the tray icon to restore')
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
            if tray_icon_initialized:
                tray_icon.notify('Speech Assistant is running in the background', 'Click the tray icon to restore')
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
            tray_icon.stop()
        
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
            shortcut.TargetPath = sys.executable
            shortcut.WorkingDirectory = os.path.dirname(sys.executable)
            shortcut.IconLocation = os.path.join(os.path.dirname(sys.executable), "needyamin.ico")
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
            shortcut.TargetPath = sys.executable
            shortcut.WorkingDirectory = os.path.dirname(sys.executable)
            shortcut.IconLocation = os.path.join(os.path.dirname(sys.executable), "needyamin.ico")
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
            shortcut.TargetPath = sys.executable
            shortcut.WorkingDirectory = os.path.dirname(sys.executable)
            shortcut.IconLocation = os.path.join(os.path.dirname(sys.executable), "needyamin.ico")
            shortcut.Save()
            print(f"Desktop shortcut created at: {desktop_shortcut_path}")
    except Exception as e:
        print(f"Error creating desktop shortcut: {e}")

# Create all shortcuts
create_startup_shortcut()
create_start_menu_shortcut()
create_desktop_shortcut()