import sys
import threading
import sounddevice as sd
import numpy as np
import speech_recognition as sr
import pyttsx3
import tkinter as tk
from tkinter import scrolledtext, ttk
import openai

# ─── CONFIG ───────────────────────────────────────────────────────
SAMPLE_RATE = 16000
CHANNELS = 1

# Paste your OpenRouter API key here
openai.api_key = "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"
openai.api_base = "https://openrouter.ai/api/v1"
# ─────────────────────────────────────────────────────────────────

recognizer = sr.Recognizer()
tts = pyttsx3.init()

recording = threading.Event()
frames = []

# Get the default input device (microphone)
default_mic_index = sd.default.device[0]  # Get the index of the default input device
all_devices = sd.query_devices()

# Filter out only input devices (microphones)
microphones = []
for device in all_devices:
    if device['max_input_channels'] > 0:  # Check if the device supports input
        microphones.append(device['name'])

# Try to set the default microphone as the first available option
selected_mic = None
for mic in microphones:
    if mic == sd.query_devices(default_mic_index)["name"]:
        selected_mic = mic
        break

# If no default microphone is found, fall back to the first available mic
if not selected_mic and microphones:
    selected_mic = microphones[0]

def audio_callback(indata, frames_count, time, status):
    if status:
        print(f"Audio status: {status}", file=sys.stderr)
    frames.append(indata.copy())

def start_recording():
    frames.clear()
    recording.set()
    start_btn.config(state=tk.DISABLED, bg='gray')  # Disable start button and change color
    stop_btn.config(state=tk.NORMAL, bg='lightcoral')  # Enable stop button and change color
    output.insert(tk.END, "🎤 Recording... Click 'Stop' when done.\n")
    output.see(tk.END)

    with sd.InputStream(samplerate=SAMPLE_RATE, channels=CHANNELS,
                        dtype='int16', callback=audio_callback):
        while recording.is_set():
            sd.sleep(100)

    process_audio()

def stop_recording():
    recording.clear()
    start_btn.config(state=tk.NORMAL, bg='lightgreen')  # Enable start button and reset color
    stop_btn.config(state=tk.DISABLED, bg='gray')  # Disable stop button and change color

def process_audio():
    if not frames:
        output.insert(tk.END, "⚠️ No audio captured.\n\n")
        output.see(tk.END)
        return

    audio_np = np.concatenate(frames, axis=0).astype(np.float32)
    if np.max(np.abs(audio_np)) == 0:
        output.insert(tk.END, "⚠️ Silence detected.\n\n")
        output.see(tk.END)
        return

    audio_np = (audio_np / np.max(np.abs(audio_np)) * 32767).astype(np.int16)
    audio_bytes = audio_np.tobytes()
    audio_data = sr.AudioData(audio_bytes, SAMPLE_RATE, 2)

    try:
        text = recognizer.recognize_google(audio_data, language="en-US")
        output.insert(tk.END, f"Recognized Text: {text}\n")
    except sr.UnknownValueError:
        output.insert(tk.END, "⚠️ Could not understand audio. Please type below.\n\n")
        output.see(tk.END)
        manual_input_frame.grid(row=2, column=0, padx=10, pady=10, sticky='ew')  # Show manual input frame
        return
    except sr.RequestError as e:
        output.insert(tk.END, f"⚠️ API error: {e}\n\n")
        output.see(tk.END)
        return
    except Exception as e:
        output.insert(tk.END, f"⚠️ Error: {e}\n\n")
        output.see(tk.END)
        return

    show_correction(text)

def submit_manual_text():
    text = manual_entry.get()
    manual_entry.delete(0, tk.END)
    manual_input_frame.grid_forget()  # Hide manual input frame
    show_correction(text)

def show_correction(text):
    corrected = correct_grammar_gpt(text)
    output.insert(tk.END, f"You said:    {text}\n")
    output.insert(tk.END, f"Corrected:   {corrected}\n\n")
    output.see(tk.END)

    tts.say(corrected)
    tts.runAndWait()

def correct_grammar_gpt(text):
    try:
        response = openai.ChatCompletion.create(
            model="openai/gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an English grammar assistant. Fix grammar and make the sentence natural."},
                {"role": "user", "content": text}
            ]
        )
        return response['choices'][0]['message']['content'].strip()
    except Exception as e:
        return f"(Error: {e})"

def select_microphone():
    global selected_mic
    selected_mic = mic_dropdown.get()
    print(f"Selected microphone: {selected_mic}")

# ─── GUI SETUP ────────────────────────────────────────────────────
root = tk.Tk()
root.title("🎙️ Speech → 🧠 AI Grammar Fix → 🗣️ Voice")
root.geometry("620x460")
root.resizable(True, True)

output = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Arial", 12))
output.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

# Configure grid expansion
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

btn_frame = tk.Frame(root)
btn_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky='ew')

start_btn = tk.Button(btn_frame, text="Start Recording", font=("Arial", 12),
                      command=lambda: threading.Thread(target=start_recording, daemon=True).start(), bg='lightgreen')
stop_btn = tk.Button(btn_frame, text="Stop Recording", font=("Arial", 12),
                     command=stop_recording, state=tk.DISABLED, bg='gray')

start_btn.grid(row=0, column=0, padx=5, sticky='ew')
stop_btn.grid(row=0, column=1, padx=5, sticky='ew')

# Dropdown for microphone selection
mic_label = tk.Label(root, text="Select Microphone:", font=("Arial", 12))
mic_label.grid(row=2, column=0, padx=10, pady=5, sticky='w')

mic_dropdown = ttk.Combobox(root, values=microphones, state="readonly", font=("Arial", 12))
mic_dropdown.set(selected_mic if selected_mic else "No microphones detected")  # Set the default microphone
mic_dropdown.grid(row=2, column=1, padx=10, pady=5, sticky='ew')
mic_dropdown.bind("<<ComboboxSelected>>", lambda e: select_microphone())

manual_input_frame = tk.Frame(root)
manual_entry = tk.Entry(manual_input_frame, font=("Arial", 12), width=40)
submit_btn = tk.Button(manual_input_frame, text="Submit", font=("Arial", 12), command=submit_manual_text)

manual_entry.grid(row=0, column=0, padx=5)
submit_btn.grid(row=0, column=1, padx=5)

root.grid_rowconfigure(2, weight=0)
root.grid_columnconfigure(1, weight=1)

root.mainloop()
