# 🧠 Tklish – Your AI Grammar & Speech Assistant

**Tklish** is a smart desktop application that combines **real-time grammar correction**, **pronunciation feedback**, and **AI-powered conversation practice** in a sleek, modern GUI. Designed for language learners, professionals, and anyone aiming to improve their English communication skills.

> 🌐 Works both **online (GPT-3.5)** and **offline** with built-in grammar correction rules.
---

![Image](https://github.com/user-attachments/assets/556d3502-fcd7-41db-b13c-f9945dca80da)


## ✨ Key Features

### 🎙️ Speech Processing
- **Real-time speech recognition** using `speech_recognition`
- **Text-to-speech** conversion via `pyttsx3`
- **Audio recording/playback** with `sounddevice` + `numpy`

### 🤖 AI Integration
- Grammar correction powered by **OpenAI GPT-3.5 Turbo**
- Custom API integration using **OpenRouter.ai**
- Online/Offline modes with API key management

### 🖥️ User Interface
- **Modern chat-style interface** with message bubbles & timestamps
- **System tray support** via `pystray`
- **Copy-to-clipboard**, **loading animations**, and status indicators

### ⚙️ System Integration
- **Windows startup** and background running
- **Start Menu/Desktop shortcut creation**
- **Microphone monitoring & internet status checks**

### 🔐 Security & Storage
- Password-protected settings
- API key stored securely in a local `SQLite` database
- Secure API communication with error logging

---

## 🧰 Tech Stack

| Area | Libraries/Technologies |
|------|------------------------|
| GUI | `tkinter`, `ttk`, `PIL`, `pystray` |
| Speech | `speech_recognition`, `pyttsx3`, `sounddevice`, `numpy` |
| AI/API | `openai`, `OpenRouter.ai` |
| System | `win32gui`, `win32con`, `win32com.client`, `winshell`, `pythoncom` |
| Database | `SQLite` for local persistent storage |

---

## 💻 System Requirements

### ✅ Windows Version
- Compatible with **Windows 10 / 11**

### 📦 Minimum
- 4GB RAM
- 2GHz Processor
- 500MB Storage

### 🚀 Recommended
- 8GB RAM
- 2.5GHz+ Processor
- 1GB Storage

---

## 📦 Installation

> Coming soon: Full installer and setup guide.

For now, clone the repo and run with Python:

```bash
git clone https://github.com/needyamin/speak-english-virtual-assistant.git
cd speak-english-virtual-assistant
pip install -r requirements.txt
python main.py


