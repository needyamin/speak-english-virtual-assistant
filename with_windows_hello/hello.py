import ctypes
from ctypes import wintypes
import tkinter as tk

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

def authenticate():
    # Initialize CREDUI_INFO structure
    ui_info = CREDUI_INFO()
    ui_info.cbSize = ctypes.sizeof(CREDUI_INFO)
    ui_info.hwndParent = wintypes.HWND(root.winfo_id())
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
        0x00000001                    # CREDUIWIN_GENERIC flag
    )

    if result == 0:
        print("Authentication successful!")
        # Only print password
        print("Password:", password.value)
        
        # Here you would typically validate credentials
        # For demonstration, just clear the password buffer
        ctypes.memset(password, 0, ctypes.sizeof(password))
    else:
        print("Authentication failed or canceled. Error code:", result)

# Create Tkinter UI
root = tk.Tk()
root.title("Windows Password Only Auth Demo")

auth_button = tk.Button(
    root,
    text="Authenticate with Password",
    command=authenticate,
    padx=20,
    pady=10
)
auth_button.pack(padx=50, pady=50)

root.mainloop()