import threading
import time
import random
import subprocess
import ctypes

import keyboard
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk

from english_words import get_english_words_set

import sys, os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


class LastLetterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Last Letter Helper")
        self.root.resizable(False, False)
        self.root.attributes('-topmost', True)

        try:
            self.root.iconbitmap(resource_path("LastLetter.ico"))
        except Exception as e:
            print("Icon load error:", e)

        self.wordlist: list[str] = []
        self.wordlist_loaded = False
        self.wordlist_error = None

        self.used_words: set[str] = set()

        self.prefix_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="Random Words")

        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)

        label = tk.Label(main_frame, text="Starting letters:")
        label.grid(row=0, column=0, sticky="w")

        self.entry = tk.Entry(main_frame, textvariable=self.prefix_var, width=20)
        self.entry.grid(row=1, column=0, columnspan=2, sticky="we", pady=(2, 6))
        self.entry.focus_set()
        self.entry.bind("<Control-Return>", self.on_ctrl_enter)

        self.status_var = tk.StringVar(value="Loading word list...")
        status_label = tk.Label(main_frame, textvariable=self.status_var, fg="gray")
        status_label.grid(row=2, column=0, sticky="w", pady=(0, 6))

        self.roblox_status_var = tk.StringVar(value="Roblox not Found")
        self.roblox_status_label = tk.Label(main_frame, textvariable=self.roblox_status_var, fg="red")
        self.roblox_status_label.grid(row=2, column=1, sticky="e", pady=(0, 6))

        start_button = tk.Button(main_frame, text="Play round", command=self.on_play_round)
        start_button.grid(row=3, column=0, sticky="we", pady=(0, 4))

        view_used_button = tk.Button(main_frame, text="View Used Words", command=self.show_used_words)
        view_used_button.grid(row=3, column=1, sticky="we", pady=(0, 4), padx=(4, 0))

        speed_label = tk.Label(main_frame, text="Typing speed (ms per char):")
        speed_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self.speed_var = tk.DoubleVar(value=30.0)
        speed_scale = tk.Scale(
            main_frame,
            from_=10,
            to=200,
            orient="horizontal",
            variable=self.speed_var,
            showvalue=True,
            length=200,
            resolution=5,
        )
        speed_scale.grid(row=5, column=0, columnspan=2, sticky="we", pady=(0, 4))

        style = ttk.Style()
        style.configure('TCombobox', 
                       padding=4, 
                       relief='raised', 
                       background='#f0f0f0',
                       arrowcolor='#333333')
        
        mode_frame = ttk.Frame(main_frame)
        mode_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=(4, 0))
        
        mode_label = ttk.Label(mode_frame, text="Word selection mode:")
        mode_label.pack(side='left', padx=(0, 10))
        
        self.mode_combobox = ttk.Combobox(
            mode_frame,
            textvariable=self.mode_var,
            values=["Random Words", "Short Words", "Long Words"],
            state='readonly',
            width=15,
            style='TCombobox'
        )
        self.mode_combobox.current(0)
        self.mode_combobox.pack(side='left', fill='x', expand=True)
        
        mode_frame.configure(style='TFrame')
        style.configure('TFrame', 
                       borderwidth=1, 
                       relief='solid',
                       padding=4)

        clear_words_button = tk.Button(main_frame, text="Clear Used Words", command=self.on_clear_cache)
        clear_words_button.grid(row=7, column=0, sticky="we", pady=(0, 4))

        quit_button = tk.Button(main_frame, text="Quit", command=self.root.destroy)
        quit_button.grid(row=7, column=1, sticky="we", pady=(0, 4), padx=(4, 0))

        credit_label = tk.Label(main_frame, text="Made by elDziad0", fg="gray")
        credit_label.grid(row=8, column=0, columnspan=2, sticky="e")

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        self.root.bind("<FocusIn>", lambda event: self.entry.focus_set())

        threading.Thread(target=self.load_wordlist, daemon=True).start()
        self.refresh_roblox_status()

    def load_wordlist(self) -> None:
        try:
            words_set = get_english_words_set(["web2"], lower=False)
            self.wordlist = sorted(words_set)
            self.wordlist_loaded = True
            self.wordlist_error = None
            self.root.after(0, lambda: self.status_var.set("Word list loaded."))
        except Exception as exc:
            self.wordlist_loaded = False
            self.wordlist_error = exc
            self.root.after(0, lambda: self.status_var.set("Failed to load word list."))

    def find_completion(self, prefix: str) -> str | None:
        if not self.wordlist_loaded or not self.wordlist:
            return None

        lower_prefix = prefix.lower()

        candidates: list[str] = []
        for word in self.wordlist:
            if word in self.used_words:
                continue
            if not word.lower().startswith(lower_prefix):
                continue
            if len(word) <= len(prefix):
                continue
            candidates.append(word)

        if not candidates:
            return None

        mode = self.mode_var.get()
        if mode == "Short Words":
            chosen = min(candidates, key=len)
        elif mode == "Long Words":
            chosen = max(candidates, key=len)
        else:
            chosen = random.choice(candidates)

        self.used_words.add(chosen)
        return chosen[len(prefix) :]

    def _is_roblox_running(self) -> bool:
        try:
            startupinfo = None
            creationflags = 0
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW

            output = subprocess.check_output(
                ["tasklist"],
                text=True,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
        except Exception:
            return False
        return "RobloxPlayerBeta.exe" in output

    def _focus_roblox_window(self) -> None:
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            hwnd = user32.FindWindowW(None, "Roblox")
            if hwnd:
                if user32.IsIconic(hwnd):
                    user32.ShowWindow(hwnd, 9)
                user32.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def refresh_roblox_status(self) -> None:
        running = self._is_roblox_running()
        if running:
            self.roblox_status_var.set("Roblox Found")
            self.roblox_status_label.config(fg="green")
        else:
            self.roblox_status_var.set("Roblox not Found")
            self.roblox_status_label.config(fg="red")

    def on_clear_cache(self) -> None:
        """Clear the list of used words and update the UI if the used words window is open."""
        self.used_words.clear()
        if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
            self.update_used_words_list()

    def on_play_round(self) -> None:
        prefix = self.prefix_var.get().strip()
        if not prefix:
            messagebox.showwarning("Last Letter Helper", "Please enter starting letters.")
            return

        if not self.wordlist_loaded:
            messagebox.showwarning("Last Letter Helper", "Word list is still loading or failed. Try again in a moment.")
            return

        completion = self.find_completion(prefix)
        if completion is None:
            messagebox.showinfo("Last Letter Helper", "No word found starting with those letters.")
            return

        self.root.withdraw()
        self.prefix_var.set("")

        roblox_running = self._is_roblox_running()
        if roblox_running:
            self._focus_roblox_window()
            self.roblox_status_var.set("Roblox Found")
            self.roblox_status_label.config(fg="green")
        else:
            self.roblox_status_var.set("Roblox not Found")
            self.roblox_status_label.config(fg="red")

        threading.Thread(target=self._type_after_delay, args=(completion,), daemon=True).start()

    def on_ctrl_enter(self, event):
        self.on_play_round()
        return "break"

    def show_used_words(self) -> None:
        """Show a window with all used words, updating in real-time."""
        if not hasattr(self, 'used_words_window') or not self.used_words_window.winfo_exists():
            self.used_words_window = tk.Toplevel(self.root)
            self.used_words_window.title("Used Words")
            self.used_words_window.resizable(True, True)
            self.used_words_window.attributes('-topmost', True)
            
            frame = ttk.Frame(self.used_words_window)
            frame.pack(fill='both', expand=True, padx=5, pady=5)
            
            scrollbar = ttk.Scrollbar(frame)
            scrollbar.pack(side='right', fill='y')
            
            self.used_words_listbox = tk.Listbox(
                frame, 
                yscrollcommand=scrollbar.set,
                font=('Consolas', 10),
                width=30,
                height=15
            )
            self.used_words_listbox.pack(side='left', fill='both', expand=True)
            
            scrollbar.config(command=self.used_words_listbox.yview)
            
            self.used_words_count = tk.Label(
                self.used_words_window, 
                text=f"Words used: {len(self.used_words)}",
                anchor='w'
            )
            self.used_words_count.pack(fill='x', padx=5, pady=(0, 5))
            
            close_button = ttk.Button(
                self.used_words_window, 
                text="Close", 
                command=self.used_words_window.destroy
            )
            close_button.pack(pady=(0, 5))
            
            self.used_words_window.update_idletasks()
            width = self.used_words_window.winfo_width()
            height = self.used_words_window.winfo_height()
            x = (self.used_words_window.winfo_screenwidth() // 2) - (width // 2)
            y = (self.used_words_window.winfo_screenheight() // 2) - (height // 2)
            self.used_words_window.geometry(f'+{x}+{y}')
            
            self.used_words_window.protocol("WM_DELETE_WINDOW", self.used_words_window.destroy)
        
        self.update_used_words_list()
        
        self.used_words_window.lift()
        self.used_words_window.focus_force()
    
    def update_used_words_list(self) -> None:
        """Update the used words listbox with current used words."""
        if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
            self.used_words_listbox.delete(0, tk.END)
            
            for word in sorted(self.used_words, key=str.lower):
                self.used_words_listbox.insert(tk.END, word)
            
            self.used_words_count.config(text=f"Words used: {len(self.used_words)}")
            
            self.used_words_window.after(500, self.update_used_words_list)
    
    def _type_after_delay(self, completion: str) -> None:
        time.sleep(1.0)
        try:
            delay = max(0.005, float(self.speed_var.get()) / 1000.0)
            for ch in completion:
                keyboard.press_and_release(ch)
                time.sleep(delay)
            keyboard.send("enter")
            time.sleep(1.0)
        finally:
            self.root.after(0, self.root.deiconify)
            
            if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
                self.update_used_words_list()


def main() -> None:
    root = tk.Tk()
    app = LastLetterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()