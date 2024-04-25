import ctypes
from win32api import GetModuleHandle
from win32con import IMAGE_ICON, LR_DEFAULTSIZE, LR_LOADFROMFILE, WM_DESTROY, WM_COMMAND, TPM_LEFTALIGN, TPM_RIGHTBUTTON, WM_USER, WM_QUIT
from win32gui import (CreateWindow, DefWindowProc, DestroyWindow, LoadImage, NIM_ADD, NIM_DELETE, NIF_ICON, NIF_MESSAGE, NIF_TIP, PostQuitMessage, Shell_NotifyIcon, UpdateWindow, WNDCLASS, TrackPopupMenu, CreatePopupMenu, AppendMenu, NIF_INFO, NIM_MODIFY)
import win32gui
import win32con
import win32clipboard
from langchain_openai import ChatOpenAI
from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import scrolledtext, font
import time
import win32api

path="C:/path_to/API_TOKEN.txt"

OPENAI_API_KEY=open(path, 'r').read()
os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY
llm = ChatOpenAI(model='gpt-3.5-turbo-0125')

check_state_option = False
check_state_notifications = False

# Constants for simple hotkeys
ID_HOTKEY_R = 193
ID_HOTKEY_C = 2900
VK_R = 0x52  # Virtual-Key code for 'R'
VK_C = 0x43  # Virtual-Key code for 'C'

HOTKEYS = {
    ID_HOTKEY_R: (VK_R, win32con.MOD_ALT),  # Alt + R
    ID_HOTKEY_C: (VK_C, win32con.MOD_ALT | win32con.MOD_SHIFT)  # Alt + Shift + C
}





try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Set DPI awareness to system DPI
except AttributeError:
    pass

def alt_r():
    set_clipboard_text(rewrite_text(get_clipboard_text()))

def alt_shift_c():
    set_clipboard_text(correct_text(get_clipboard_text()))

def register_hotkeys(hwnd):
    global HOTKEYS
    for id, (vk, modifiers) in HOTKEYS.items():
        win32gui.RegisterHotKey(hwnd, id, modifiers, vk)



# Chat functionality
def send_message(event=None):
    """Function to handle sending messages."""
    user_input = entry_field.get()
    if user_input:
        # Update the chat display
        chat_display.config(state=tk.NORMAL)  # Enable editing of the Text widget
        chat_display.insert(tk.END, f"You: {user_input}\n", 'user')
        entire_chat = chat_display.get("1.0", tk.END)  # Get all text from the widget
        chat_display.config(state=tk.DISABLED)
        template = """You are 'ChatBot' ready to help the user 'You'.
        You:I will talk to you. {question}

        ChatBot:"""
        answer=llm_interaction(template, entire_chat)
        # Here you would typically add code to send the user_input to your backend (like a language model) and get the response
        # For demonstration, we'll just echo the input back as a simulated response
        
        answer = answer['text'].strip()
        simulated_response = answer
        chat_display.config(state=tk.NORMAL)  # Enable editing of the Text widget
        chat_display.insert(tk.END, f"ChatBot: {simulated_response}\n", 'bot')
        chat_display.config(state=tk.DISABLED)
        chat_display.see(tk.END)  # Scroll to the bottom of the chat display
        entry_field.delete(0, tk.END)  # Clear the input field

def setup_ui():
    """Setup the chat UI components."""
    global entry_field, chat_display
    root = tk.Tk()
    root.title("Chat")
    root.resizable(True, True)  # Allow window resizing
    # Chat display area
    chat_display = scrolledtext.ScrolledText(root, state='disabled', height=15, width=50,
                                             wrap=tk.WORD, font=('Helvetica', 10), padx=5, pady=5)
    chat_display.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Define text alignment tags
    chat_display.tag_configure('user', justify='right', background='#DCF8C6', relief='raised',
                                lmargin1=50, lmargin2=50, rmargin=10)
    chat_display.tag_configure('bot', justify='left', background='#FFFFFF', relief='raised',
                                lmargin1=10, lmargin2=10, rmargin=50)
    # Frame for the entry field and send button
    entry_frame = tk.Frame(root)
    entry_frame.pack(padx=10, pady=10, fill=tk.X)

    # Entry field
    entry_field = tk.Entry(entry_frame, width=40, font=('Helvetica', 10))
    entry_field.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)

    # Send button
    entry_field.bind('<Return>', send_message)  # Bind the Enter key to the send_message function
    send_button = tk.Button(entry_frame, text="Send", command=send_message)
    send_button.pack(side=tk.RIGHT)

    root.mainloop()

# ---------------------------

def llm_interaction(template, question):
    global llm
    prompt_template = PromptTemplate.from_template(template)
    answer_chain = LLMChain(llm=llm, prompt=prompt_template)
    answer = answer_chain.invoke(question)
    return answer

def show_text_window(text):
    """ Create a text display window using tkinter, positioned in the top right corner of the screen. """
    root = tk.Tk()
    root.title("Text Output")

    # Set a reasonable window size
    window_width = 400
    window_height = 200

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate position for the top right corner
    x = screen_width - window_width - 10  # 10 pixels padding from the screen edge
    y = 10  # 10 pixels padding from the top

    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Create a scrollbar and attach it to a Text widget
    scrollbar = tk.Scrollbar(root)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Create a Text widget with a scrollbar
    text_widget = tk.Text(root, height=10, width=50, yscrollcommand=scrollbar.set, wrap=tk.WORD, font=("Helvetica", 12))
    text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
    text_widget.insert(tk.END, text)

    # Link scrollbar to the text widget
    scrollbar.config(command=text_widget.yview)

    # Disable editing of the text
    text_widget.config(state=tk.DISABLED)

    # Start the GUI loop
    root.mainloop()

def get_clipboard_text():
    """ Retrieves text from the clipboard and converts bytes to a string if necessary. """
    win32clipboard.OpenClipboard()
    text = ''
    try:
        # Use CF_UNICODETEXT to get text as a string (handles UTF-16 text)
        text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
    except TypeError:
        text = ''  # Clipboard does not contain text
    finally:
        win32clipboard.CloseClipboard()
    return text

def set_clipboard_text(text):
    """ Inserts text into the clipboard. """
    win32clipboard.OpenClipboard()  # Open the clipboard to start clipboard operations.
    win32clipboard.EmptyClipboard()  # Clear the clipboard of its current content.
    try:
        # Use CF_UNICODETEXT to handle Unicode text
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()  # Always close the clipboard after operations are completed.

def show_notification(hwnd, title, msg):
    """Display a balloon notification."""
    flags = NIF_INFO
    nid = (hwnd, 0, flags, 0, None, "Tooltip", msg, 200, title)
    win32gui.Shell_NotifyIcon(NIM_MODIFY, nid)



def correct_text(text):
    global check_state_notifications
    question = text
    set_clipboard_text('Waiting...')

    template = """Correct just the grammar in the following text: {question}

    Answer:"""
    if check_state_notifications:
        show_notification(hwnd, "Processing", "Sending correction request to OpenAI...")
    answer=llm_interaction(template, question)
    if check_state_notifications:
        show_notification(hwnd, "Completed", "Text correction completed.")
    answer = answer['text'].strip()
    
    return answer

def rewrite_text(text):
    global check_state_notifications
    question = text
    set_clipboard_text('Waiting...')

    template = """Rewrite better the following text: {question}

    Answer:"""
    if check_state_notifications:
        show_notification(hwnd, "Processing", "Sending rewrite request to OpenAI...")
    answer=llm_interaction(template, question)
    if check_state_notifications:
        show_notification(hwnd, "Completed", "Text rewriting completed.")
    answer = answer['text'].strip()
    
    return answer

def summarize_text(text):
    global check_state_notifications
    question = text
    set_clipboard_text('Waiting...')

    template = """Summarize the following text: {question}

    Answer:"""
    if check_state_notifications:
        show_notification(hwnd, "Processing", "Sending summarization request to OpenAI...")
    answer=llm_interaction(template, question)
    if check_state_notifications:
        show_notification(hwnd, "Completed", "Text summarization completed.")
    answer = answer['text'].strip()
    
    return answer

def answer_text(text):
    global check_state_notifications
    question = text
    set_clipboard_text('Waiting...')

    template = """Answer the following question: {question}

    Answer:"""
    if check_state_notifications:
        show_notification(hwnd, "Processing", "Sending question answering request to OpenAI...")
    answer=llm_interaction(template, question)
    if check_state_notifications:
        show_notification(hwnd, "Completed", "Question answering completed.")
    answer = answer['text'].strip()
    
    return answer

def save_text_to_file(file_path):
    """Function to save text from the text widget to a file."""
    try:
        # Get text from text widget
        text_to_save = text_widget.get("1.0", tk.END)
        # Save text to the specified file
        with open(file_path, 'w') as file:
            file.write(text_to_save)
        messagebox.showinfo("Success", "Text saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save text: {str(e)}")

def token():
    # Create the main window
    root = tk.Tk()
    root.title("Token")
    root.geometry("400x250")  # Width x Height

    # Create a Text widget
    global text_widget
    text_widget = tk.Text(root, height=10, width=48)
    text_widget.pack(padx=10, pady=10)

    # Create a Save Button
    save_button = tk.Button(root, text="Save Token", command=lambda: save_text_to_file(path))
    save_button.pack(pady=10)

    # Start the GUI event loop
    root.mainloop()

def wndproc(hwnd, msg, wparam, lparam):
    global check_state_option
    global check_state_notifications

    if msg == win32con.WM_DESTROY:
        nid = (hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        #unregister_hotkeys()
        PostQuitMessage(0)  # Terminate the app.
    elif msg == win32con.WM_USER + 20:  # Custom message for tray icon interaction
        if lparam == win32con.WM_RBUTTONUP:  # Right-click release
            show_exit_menu(hwnd)
        elif lparam == win32con.WM_LBUTTONUP:  # Left-click release
            print("Icon clicked!")
            show_main_menu(hwnd)
    elif msg == WM_COMMAND:
        if wparam == 1000:  # ID for the "Exit" command
            win32gui.PostMessage(hwnd, WM_DESTROY, 0, 0)
        elif wparam == 1001:  # ID for the "Rewrite" command
            set_clipboard_text(rewrite_text(get_clipboard_text()))
            if check_state_option:
                show_text_window(get_clipboard_text())
        elif wparam == 1002:  # ID for the "Summarize" command
            set_clipboard_text(summarize_text(get_clipboard_text()))
            if check_state_option:
                show_text_window(get_clipboard_text())
        elif wparam == 1003:  # ID for the "Answer" command
            set_clipboard_text(answer_text(get_clipboard_text()))
            if check_state_option:
                show_text_window(get_clipboard_text())
        elif wparam == 1004:  # ID for the "Correct Grammar" command
            set_clipboard_text(correct_text(get_clipboard_text()))
            if check_state_option:
                show_text_window(get_clipboard_text())
        elif wparam == 1005:  # ID for the "Option" command
            check_state_option = not check_state_option
        elif wparam == 1006:  # ID for the "Token" command
            token()
        elif wparam == 1007:  # ID for the "Chat" command
            setup_ui()
        elif wparam == 1008:  # ID for the "Notifications" command
            check_state_notifications = not check_state_notifications
    elif msg == win32con.WM_HOTKEY:
        if wparam == ID_HOTKEY_R:
            alt_r()
        elif wparam == ID_HOTKEY_C:
            alt_shift_c()
    return DefWindowProc(hwnd, msg, wparam, lparam)

def show_exit_menu(hwnd):
    """Create and display an exit menu with a checkbox."""
    global check_state_option

    menu = win32gui.CreatePopupMenu()
    # Add a checkable menu item
    if check_state_option:
        flag = win32con.MF_STRING | win32con.MF_CHECKED
    else:
        flag = win32con.MF_STRING | win32con.MF_UNCHECKED

    if check_state_notifications:
        flag_notif = win32con.MF_STRING | win32con.MF_CHECKED
    else:
        flag_notif = win32con.MF_STRING | win32con.MF_UNCHECKED

    win32gui.AppendMenu(menu, flag_notif, 1008, "Notifications")
    win32gui.AppendMenu(menu, flag, 1005, "Window")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1006, "Token")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1000, "Exit")
    pos = win32gui.GetCursorPos()
    win32gui.SetForegroundWindow(hwnd)
    win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN | win32con.TPM_RIGHTBUTTON, pos[0], pos[1], 0, hwnd, None)
    win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)

def show_main_menu(hwnd):
    menu = win32gui.CreatePopupMenu()
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1001, "Rewrite")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1002, "Summarize")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1003, "Answer Question")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1004, "Correct Grammar")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1007, "Chat")
    pos = win32gui.GetCursorPos()
    win32gui.SetForegroundWindow(hwnd)
    win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN | win32con.TPM_RIGHTBUTTON, pos[0], pos[1], 0, hwnd, None)
    win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)

def create_tray_icon(hwnd, hinst, icon_path):
    hicon = win32gui.LoadImage(hinst, icon_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
    
    if hicon == 0:  # If LoadImage failed, hicon will be 0.
        print("Could not load icon:", ctypes.GetLastError())
        return None

    flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
    nid = (hwnd, 0, flags, win32con.WM_USER+20, hicon, "ChatGPT")
    win32gui.Shell_NotifyIcon(NIM_ADD, nid)
    return nid

if __name__ == '__main__':
    hinst = GetModuleHandle(None)
    wc = WNDCLASS()
    wc.hInstance = hinst
    wc.lpszClassName = "PythonTaskbarDemo"
    wc.lpfnWndProc = wndproc
    class_atom = win32gui.RegisterClass(wc)
    hwnd = CreateWindow(class_atom, "Taskbar Demo", 0, 0, 0, 0, 0, 0, 0, hinst, None)
    UpdateWindow(hwnd)
    register_hotkeys(hwnd)
    icon_path = "path_to/Icons/oie_jpg.ico"
    nid = create_tray_icon(hwnd, hinst, icon_path)
    
    win32gui.PumpMessages()
