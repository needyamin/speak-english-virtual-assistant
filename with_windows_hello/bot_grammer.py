import sys
import threading
import sounddevice as sd
import numpy as np
import speech_recognition as sr
import pyttsx3
import tkinter as tk
from tkinter import scrolledtext, ttk, Menu, Toplevel, Label, Entry, Button
import sqlite3
import openai
import time
import os
from PIL import Image, ImageTk
import traceback
import pystray
from pystray import MenuItem as item
import ctypes
from ctypes import wintypes

# ─── CONFIG STORAGE ─────────────────────────────────────────────
DB_PATH = os.path.join(os.path.dirname(__file__), 'data.sqlite')
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()
cur.execute('''CREATE TABLE IF NOT EXISTS config (
    id INTEGER PRIMARY KEY,
    api_key TEXT,
    api_base TEXT
)''')
cur.execute('SELECT api_key, api_base FROM config WHERE id=1')
row = cur.fetchone()
if row:
    openai.api_key, openai.api_base = row
else:
    openai.api_key = ''
    openai.api_base = ''
    cur.execute('INSERT INTO config (id, api_key, api_base) VALUES (1, ?, ?)', ('', ''))
    conn.commit()

# Add a password column to the config table if it doesn't exist
cur.execute('''PRAGMA table_info(config)''')
columns = [col[1] for col in cur.fetchall()]
if 'password' not in columns:
    cur.execute('ALTER TABLE config ADD COLUMN password TEXT')
    cur.execute('UPDATE config SET password=? WHERE id=1', ('admin',))  # Default password
    conn.commit()

# Password verification
def verify_password():
    def login():
        entered_password = password_entry.get().strip()
        cur.execute('SELECT password FROM config WHERE id=1')
        stored_password = cur.fetchone()[0]
        if entered_password == stored_password:
            log("Login successful.")
            login_win.destroy()
            open_settings_interface()
        else:
            log("Incorrect password. Please try again.", level='ERROR')

    login_win = Toplevel(root)
    login_win.title("Login")
    Label(login_win, text="Enter Password:").grid(row=0, column=0, padx=5, pady=5)
    password_entry = Entry(login_win, show="*", width=30)
    password_entry.grid(row=0, column=1, padx=5, pady=5)
    Button(login_win, text="Login", command=login).grid(row=1, column=0, columnspan=2, pady=10)

# Open settings interface with password reset functionality
def open_settings_interface():
    def save_settings():
        key = api_key_entry.get().strip()
        base = api_base_entry.get().strip()
        new_password = new_password_entry.get().strip()
        confirm_password = confirm_password_entry.get().strip()

        if new_password and new_password != confirm_password:
            log("Passwords do not match. Please try again.", level='ERROR')
            return

        with sqlite3.connect(DB_PATH) as thread_conn:
            thread_cur = thread_conn.cursor()
            thread_cur.execute('UPDATE config SET api_key=?, api_base=? WHERE id=1', (key, base))
            if new_password:
                thread_cur.execute('UPDATE config SET password=? WHERE id=1', (new_password,))
            thread_conn.commit()

        openai.api_key = key
        openai.api_base = base
        log("Settings saved.")
        settings_win.destroy()

    settings_win = Toplevel(root)
    settings_win.title("Settings")
    Label(settings_win, text="OpenAI API Key:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
    api_key_entry = Entry(settings_win, width=50)
    api_key_entry.grid(row=0, column=1, padx=5, pady=5)
    api_key_entry.insert(0, openai.api_key)

    Label(settings_win, text="API Base URL:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
    api_base_entry = Entry(settings_win, width=50)
    api_base_entry.grid(row=1, column=1, padx=5, pady=5)
    api_base_entry.insert(0, openai.api_base)

    Label(settings_win, text="New Password:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
    new_password_entry = Entry(settings_win, show="*", width=50)
    new_password_entry.grid(row=2, column=1, padx=5, pady=5)

    Label(settings_win, text="Confirm Password:").grid(row=3, column=0, sticky='w', padx=5, pady=5)
    confirm_password_entry = Entry(settings_win, show="*", width=50)
    confirm_password_entry.grid(row=3, column=1, padx=5, pady=5)

    Button(settings_win, text="Save", command=save_settings).grid(row=4, column=0, columnspan=2, pady=10)

# Update open_settings to use the new password system
def open_settings():
    verify_password()

# ─────────────────────────────────────────────────

recognizer = sr.Recognizer()
tts = pyttsx3.init()
recording = threading.Event()
frames = []
stream = None

# ─── GUI SETUP ─────────────────────────────────────────
root = tk.Tk()
root.title("🎧 Speech → 🧠 AI Grammar Fix → 🗣️ Voice")
root.geometry("700x650")
ico_path = os.path.join(os.path.dirname(__file__), 'needyamin.ico')
if os.path.exists(ico_path):
    try:
        root.iconbitmap(ico_path)
    except Exception:
        pass

# Initialize logo_photo globally
logo_photo = None

# Display logo image (converted from .ico if needed)
logo_img_path = os.path.join(os.path.dirname(__file__), 'logo_display.png')
if os.path.exists(ico_path) and not os.path.exists(logo_img_path):
    try:
        img = Image.open(ico_path)
        img.save(logo_img_path, format='PNG')
    except Exception:
        logo_img_path = None

if logo_img_path and os.path.exists(logo_img_path):
    try:
        logo_img = Image.open(logo_img_path)
        logo_img = logo_img.resize((48, 48), Image.ANTIALIAS)
        logo_photo = ImageTk.PhotoImage(logo_img)  # Assign to global variable
        logo_label = tk.Label(root, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
    except Exception:
        logo_photo = None  # Ensure logo_photo is None if an error occurs

# Load icon image for logging
icon_image = None
if os.path.exists(ico_path):
    try:
        img = Image.open(ico_path)
        img = img.resize((16, 16), Image.ANTIALIAS)
        icon_image = ImageTk.PhotoImage(img)
    except Exception:
        icon_image = None

# Update the log function to ensure the image is displayed correctly
def log(message, level='INFO', icon=False):
    output.configure(state=tk.NORMAL)
    if message == 'Clear Logs':
        output.delete('1.0', tk.END)
        output.configure(state=tk.DISABLED)
        return
    timestamp = time.strftime('%d %B %Y: %I:%M %p').lstrip('0')
    output.insert(tk.END, '   ' + timestamp + '\n', 'timestamp')
    if icon and logo_photo:  # Ensure logo_photo is available
        output.image_create(tk.END, image=logo_photo)  # Insert the image
        output.insert(tk.END, '  ', level)  # Add spacing after the image
    output.insert(tk.END, message + '\n', level)
    output.configure(state=tk.DISABLED)
    output.see(tk.END)

# Microphone setup
all_devices = sd.query_devices()
microphones = [dev['name'] for dev in all_devices if dev['max_input_channels'] > 0]
selected_mic = microphones[0] if microphones else None
SAMPLE_RATE = 16000
CHANNELS = 1

# Audio callback
def audio_callback(indata, frames_count, time_info, status):
    if status:
        log(f"Audio status: {status}", level='WARNING')
    frames.append(indata.copy())

# Recording control
def start_recording():
    global stream
    frames.clear()
    recording.set()
    start_btn.config(state=tk.DISABLED)
    stop_btn.config(state=tk.NORMAL)
    log("Recording started... Click 'Stop' when done.")

    def run_stream():
        global stream
        try:
            stream = sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=CHANNELS,
                device=microphones.index(selected_mic) if selected_mic else None,
                callback=audio_callback
            )
            with stream:
                while recording.is_set():
                    sd.sleep(100)
        except Exception as e:
            log(f"Error starting recording: {e}", level='ERROR')
        finally:
            stream = None
            stop_btn.config(state=tk.DISABLED)
            start_btn.config(state=tk.NORMAL)

    threading.Thread(target=run_stream, daemon=True).start()

def stop_recording():
    recording.clear()
    log("Recording stopped.")
    if stream is None:
        start_btn.config(state=tk.NORMAL)
        stop_btn.config(state=tk.DISABLED)
    threading.Thread(target=process_audio, daemon=True).start()

# Audio processing
def process_audio():
    if not frames:
        log("No audio captured.", level='WARNING')
        return
    audio_np = np.concatenate(frames, axis=0).astype(np.float32)
    if np.max(np.abs(audio_np)) == 0:
        log("Silence detected.", level='WARNING')
        return
    audio_np = (audio_np / np.max(np.abs(audio_np)) * 32767).astype(np.int16)
    audio_bytes = audio_np.tobytes()
    audio_data = sr.AudioData(audio_bytes, SAMPLE_RATE, 2)
    try:
        text = recognizer.recognize_google(audio_data, language="en-US")
        log(f"Recognized text: {text}")
    except sr.UnknownValueError:
        log("Could not understand audio. Prompting for manual input.", level='WARNING')
        manual_input_frame.grid(row=2, column=0, columnspan=6, padx=10, pady=5, sticky='ew')
        return
    except sr.RequestError as e:
        log(f"Recognition API error: {e}", level='ERROR')
        return
    show_correction(text)

# Manual entry
def submit_manual_text():
    text = manual_entry.get("1.0", tk.END).strip()
    manual_entry.delete("1.0", tk.END)
    manual_input_frame.grid_forget()
    show_correction(text)

# Grammar correction & TTS
def show_correction(text):
    log("You said:", icon=True)  # Display the image before "You said:"
    log(text)  # Display the user's input text
    corrected = correct_grammar_gpt(text)
    log("Corrected:", icon=True)  # Display the image before "Corrected:"
    log(corrected)  # Display the corrected text
    tts.say(corrected)
    tts.runAndWait()

def correct_grammar_gpt(text):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an English grammar assistant. Fix grammar and make the sentence natural."},
                {"role": "user", "content": text}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"(Error: {e})"

# Microphone selection handler
def select_microphone(event=None):
    global selected_mic
    selected_mic = mic_dropdown.get()
    log(f"Selected microphone: {selected_mic}")

# Menu bar
menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Settings", command=open_settings)
file_menu.add_command(label="Clear Logs", command=lambda: log('Clear Logs'))
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)
help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label="About", command=lambda: log("Speech-to-Text UI v1.0"))
menu_bar.add_cascade(label="Help", menu=help_menu)
root.config(menu=menu_bar)

# Log output widget
output = scrolledtext.ScrolledText(root, wrap=tk.WORD, state=tk.DISABLED,
                                   font=("Consolas", 11), bg='#f9f9f9')
output.grid(row=0, column=1, columnspan=5, padx=10, pady=10, sticky='nsew')
for lvl, color in [('INFO', 'black'), ('WARNING', 'orange'), ('ERROR', 'red')]:
    output.tag_config(lvl, foreground=color)
output.tag_config('timestamp', foreground='gray', font=("Consolas", 9))
output.icon_image = icon_image  # Prevent garbage collection

# Control buttons and mic dropdown
btn_frame = tk.Frame(root)
btn_frame.grid(row=1, column=0, columnspan=6, pady=5, sticky='ew')
btn_frame.grid_columnconfigure(6, weight=1)
start_btn = tk.Button(btn_frame, text="Start Recording", font=("Arial", 12), bg='lightgreen',
                      command=lambda: threading.Thread(target=start_recording, daemon=True).start())
stop_btn = tk.Button(btn_frame, text="Stop Recording", font=("Arial", 12), bg='lightcoral', state=tk.DISABLED,
                     command=stop_recording)
clear_btn = tk.Button(btn_frame, text="Clear Logs", font=("Arial", 12), command=lambda: log('Clear Logs'))
settings_btn = tk.Button(btn_frame, text="Settings", font=("Arial", 12), command=open_settings)
start_btn.grid(row=0, column=0, padx=5)
stop_btn.grid(row=0, column=1, padx=5)
clear_btn.grid(row=0, column=2, padx=5)
settings_btn.grid(row=0, column=3, padx=5)

# Update the microphone dropdown styling in the GUI setup
mic_dropdown = ttk.Combobox(btn_frame, values=microphones, width=20, font=("Arial", 12))
mic_dropdown.set(selected_mic)
mic_dropdown.grid(row=0, column=4, padx=5, pady=5, sticky='ew')
mic_dropdown.bind('<<ComboboxSelected>>', select_microphone)

# Adjust column weights to ensure consistent alignment
btn_frame.grid_columnconfigure(4, weight=1)

# Manual text entry
manual_input_frame = tk.Frame(root)
manual_entry = tk.Text(manual_input_frame, height=4, width=40, font=("Arial", 12))
manual_entry.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
submit_btn = tk.Button(manual_input_frame, text="Submit", font=("Arial", 12), command=submit_manual_text)
submit_btn.pack(side=tk.TOP, pady=10)

# Function to handle system tray icon actions
def on_quit(icon, item):
    root.quit()
    icon.stop()

# Initialize the system tray icon
tray_icon = None  # Declare globally to manage its lifecycle
tray_icon_initialized = False  # Flag to ensure the tray icon is initialized only once

def setup_system_tray():
    global tray_icon, tray_icon_initialized
    if tray_icon_initialized:  # Prevent multiple tray icons
        return
    if os.path.exists(ico_path):
        tray_image = Image.open(ico_path)
        menu = pystray.Menu(item('Quit', on_quit))
        tray_icon = pystray.Icon("needyamin", tray_image, "Speech Assistant", menu)
        tray_icon_initialized = True  # Mark as initialized
        threading.Thread(target=tray_icon.run, daemon=True).start()

# Ensure the icon is displayed in the taskbar and start bar
def set_taskbar_icon():
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)  # Set the taskbar icon for the window
            # Explicitly set the taskbar icon using ctypes
            hwnd = ctypes.windll.user32.GetAncestor(root.winfo_id(), 2)  # Get the top-level window handle
            if hwnd:  # Ensure hwnd is valid
                hicon = ctypes.windll.user32.LoadImageW(
                    None, ico_path, 1, 32, 32, 0x00000010  # Load the icon as 32x32
                )
                if hicon:  # Ensure hicon is valid
                    ctypes.windll.user32.SendMessageW(hwnd, 0x80, 0, hicon)  # WM_SETICON for large icon
                    ctypes.windll.user32.SendMessageW(hwnd, 0x80, 1, hicon)  # WM_SETICON for small icon
        except Exception as e:
            print(f"Error setting taskbar icon: {e}")

# Ensure the system tray icon is removed when the application is closed
def on_close():
    global tray_icon
    if tray_icon:
        tray_icon.stop()  # Stop the system tray icon
    root.destroy()  # Close the main application window

# Call the functions to set icons
set_taskbar_icon()  # Set the taskbar/start bar icon
setup_system_tray()  # Set the system tray icon

# Bind the close event to ensure the system tray icon is removed
root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()
