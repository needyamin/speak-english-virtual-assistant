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
import win32security
import win32con
import win32api
import win32cred
import traceback
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
# ─────────────────────────────────────────────────

recognizer = sr.Recognizer()
tts = pyttsx3.init()
recording = threading.Event()
frames = []
stream = None

# ─── GUI SETUP ─────────────────────────────────────────
root = tk.Tk()
root.title("🎧 Speech → 🧠 AI Grammar Fix → 🗣️ Voice")
root.geometry("700x600")
ico_path = os.path.join(os.path.dirname(__file__), 'needyamin.ico')
if os.path.exists(ico_path):
    try:
        root.iconbitmap(ico_path)
    except Exception:
        pass

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
        logo_photo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(root, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
    except Exception:
        pass

# Load icon image for logging
icon_image = None
if os.path.exists(ico_path):
    try:
        img = Image.open(ico_path)
        img = img.resize((16, 16), Image.ANTIALIAS)
        icon_image = ImageTk.PhotoImage(img)
    except Exception:
        icon_image = None

# Logging helper
def log(message, level='INFO', icon=False):
    output.configure(state=tk.NORMAL)
    if message == 'Clear Logs':
        output.delete('1.0', tk.END)
        output.configure(state=tk.DISABLED)
        return
    timestamp = time.strftime('%d %B %Y: %I:%M %p').lstrip('0')
    output.insert(tk.END, '   ' + timestamp + '\n', 'timestamp')
    if icon and output.icon_image:
        output.image_create(tk.END, image=output.icon_image)
        output.insert(tk.END, '  ', level)
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
    log(f"You said: {text}")
    corrected = correct_grammar_gpt(text)
    log(f"Corrected: {corrected}", icon=True)
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

# Define CREDUI_INFO structure
class CREDUI_INFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hwndParent", wintypes.HWND),
        ("pszMessageText", wintypes.LPCWSTR),
        ("pszCaptionText", wintypes.LPCWSTR),
        ("hbmBanner", wintypes.HBITMAP),
    ]

# Set up Windows CredUI API
credui = ctypes.WinDLL('credui.dll')

# Configure CredUIPromptForCredentialsW function signature
credui.CredUIPromptForCredentialsW.argtypes = [
    ctypes.POINTER(CREDUI_INFO),
    wintypes.LPCWSTR,
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    wintypes.ULONG,
    wintypes.LPWSTR,
    wintypes.ULONG,
    ctypes.POINTER(wintypes.BOOL),
    wintypes.DWORD,
]
credui.CredUIPromptForCredentialsW.restype = wintypes.DWORD

# Update the CredUIPromptForCredentialsW flag
CREDUIWIN_GENERIC = 0x00000001
CREDUIWIN_SECURE_PROMPT = 0x00001000  # Use this flag for secure prompts

def verify_windows_user():
    try:
        # Initialize CREDUI_INFO structure
        ui_info = CREDUI_INFO()
        ui_info.cbSize = ctypes.sizeof(CREDUI_INFO)
        ui_info.hwndParent = wintypes.HWND(root.winfo_id())  # Set parent to the Tkinter window
        ui_info.pszMessageText = "Enter your password"
        ui_info.pszCaptionText = "Password Verification"
        ui_info.hbmBanner = None

        # Username buffer is empty, password buffer for input
        username = ctypes.create_unicode_buffer("")
        password = ctypes.create_unicode_buffer(256)
        save_credentials = wintypes.BOOL(False)

        # Call Windows credential dialog
        result = credui.CredUIPromptForCredentialsW(
            ctypes.byref(ui_info),        # CREDUI_INFO
            "PasswordOnlyApp",            # Target name
            None,                         # Reserved
            0,                            # Auth error code
            username,                     # Username buffer
            0,                            # Username max length (0 disables username field)
            password,                     # Password buffer
            256,                          # Password max length
            ctypes.byref(save_credentials), # Save checkbox
            CREDUIWIN_GENERIC | CREDUIWIN_SECURE_PROMPT  # Use combined flags for better compatibility
        )

        if result == 0:
            print("Authentication successful!")
            # Only print password
            print("Password:", password.value)
            
            # Here you would typically validate credentials
            # For demonstration, just clear the password buffer
            ctypes.memset(password, 0, ctypes.sizeof(password))
            open_settings_interface()  # Open settings interface after successful authentication
            return True
        else:
            print("Authentication failed or canceled. Error code:", result)
            return False
    except Exception as e:
        # Log error to file for debugging
        with open(os.path.join(os.path.dirname(__file__), "error.log"), "a", encoding="utf-8") as f:
            f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Authentication error:\n")
            traceback.print_exc(file=f)
        log(f"Authentication failed: {e}", level='ERROR')
        return False

def open_settings_interface():
    # Open settings dialog
    def save_settings():
        key = api_key_entry.get().strip()
        base = api_base_entry.get().strip()
        with sqlite3.connect(DB_PATH) as thread_conn:
            thread_cur = thread_conn.cursor()
            thread_cur.execute('UPDATE config SET api_key=?, api_base=? WHERE id=1', (key, base))
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

    Button(settings_win, text="Save", command=save_settings).grid(row=2, column=0, columnspan=2, pady=10)

def open_settings():
    # Trigger Windows authentication when "Settings" is clicked
    if not verify_windows_user():
        log("Access to settings denied.", level='ERROR')

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

mic_dropdown = ttk.Combobox(btn_frame, values=microphones, width=30)
mic_dropdown.set(selected_mic)
mic_dropdown.grid(row=0, column=4, padx=5)
mic_dropdown.bind('<<ComboboxSelected>>', select_microphone)

# Manual text entry
manual_input_frame = tk.Frame(root)
manual_entry = tk.Text(manual_input_frame, height=4, width=70, font=("Arial", 12))
manual_entry.pack(side=tk.LEFT)
submit_btn = tk.Button(manual_input_frame, text="Submit", font=("Arial", 12), command=submit_manual_text)
submit_btn.pack(side=tk.LEFT)

root.mainloop()
