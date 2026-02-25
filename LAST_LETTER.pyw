import threading
import time
import random
import subprocess
import ctypes
import bisect
import sys
import os

import keyboard
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
from packaging import version
import requests

from english_words import get_english_words_set

CURRENT_VERSION = "2.0.0"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

class HumanTypist:
    QWERTY_NEIGHBORS = {
        'a': 'qwsz', 'b': 'vghn', 'c': 'xdfv', 'd': 'sfcxe', 'e': 'wsdr',
        'f': 'dgrcv', 'g': 'fhvbt', 'h': 'gjbnm', 'i': 'ujko', 'j': 'hknmu',
        'k': 'jlmi', 'l': 'kop', 'm': 'njk', 'n': 'bhmj', 'o': 'iklp',
        'p': 'ol', 'q': 'wa', 'r': 'edft', 's': 'awedxz', 't': 'rfgy',
        'u': 'yhij', 'v': 'cfgb', 'w': 'qeas', 'x': 'zsdc', 'y': 'tghu', 'z': 'asx'
    }

    def __init__(self, target_wpm: float, typo_chance: float = 0.02):
        self.base_delay = 12.0 / max(1.0, target_wpm)
        self.typo_chance = typo_chance

    def _get_gaussian_delay(self) -> float:
        delay = random.gauss(self.base_delay, self.base_delay * 0.3)
        return max(0.005, delay)

    def type_word(self, word: str):
        burst_counter = 0
        
        time.sleep(random.uniform(0.3, 0.7))

        for i, char in enumerate(word):
            lower_char = char.lower()
            
            if random.random() < self.typo_chance and lower_char in self.QWERTY_NEIGHBORS:
                wrong_char = random.choice(self.QWERTY_NEIGHBORS[lower_char])
                keyboard.press_and_release(wrong_char if char.islower() else wrong_char.upper())
                
                time.sleep(self._get_gaussian_delay() * random.uniform(2.0, 3.5))
                keyboard.send("backspace")
                time.sleep(self._get_gaussian_delay() * random.uniform(1.0, 2.0))
            
            keyboard.press_and_release(char)
            
            delay = self._get_gaussian_delay()
            
            if char in ".,!? ":
                delay *= random.uniform(1.5, 3.0)
                burst_counter = 0
            else:
                burst_counter += 1
            
            if burst_counter > 4:
                delay *= random.uniform(1.1, 1.4)
                if random.random() < 0.2: 
                    burst_counter = 0
                    
            time.sleep(delay)
            
        time.sleep(random.uniform(0.1, 0.3))
        keyboard.send("enter")
        time.sleep(1.0)

class WordDatabase:
    def __init__(self):
        self.wordlist: list[str] = []
        self.used_words: set[str] = set()
        self.is_loaded = False
        self.error = None

    def load_dataset(self, sets: list[str]):
        self.is_loaded = False
        try:
            words_set = get_english_words_set(sets, lower=True)
            filtered_words = [w for w in words_set if w.isalpha()]
            self.wordlist = sorted(filtered_words)
            self.is_loaded = True
            self.error = None
        except Exception as e:
            self.error = str(e)
            self.is_loaded = False

    def get_candidates(self, prefix: str) -> list[str]:
        if not self.is_loaded or not self.wordlist:
            return []
            
        prefix = prefix.lower()
        
        idx_start = bisect.bisect_left(self.wordlist, prefix)
        
        prefix_end = prefix[:-1] + chr(ord(prefix[-1]) + 1)
        idx_end = bisect.bisect_left(self.wordlist, prefix_end, lo=idx_start)
        
        return self.wordlist[idx_start:idx_end]


class LastLetterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Last Letter Pro")
        self.root.resizable(False, False)
        self.root.attributes('-topmost', True)

        try:
            self.root.iconbitmap(resource_path("LastLetter.ico"))
        except Exception:
            pass

        self.db = WordDatabase()
        
        self.prefix_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="Random Words")
        self.dataset_var = tk.StringVar(value="Combined (web2 + gcide)")
        self.status_var = tk.StringVar(value="Awaiting initialization...")
        self.roblox_status_var = tk.StringVar(value="Roblox not Found")
        self.speed_var = tk.DoubleVar(value=80.0)
        self.typo_chance_var = tk.DoubleVar(value=3.0)

        self._build_ui()
        self.root.bind("<FocusIn>", lambda event: self.entry.focus_set())
        
        self.on_dataset_changed()
        self.refresh_roblox_status()

    def _build_ui(self):
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="Dataset:").grid(row=0, column=0, sticky="w")
        ds_combo = ttk.Combobox(
            main_frame, 
            textvariable=self.dataset_var,
            values=["web2", "gcide", "Combined (web2 + gcide)"],
            state='readonly'
        )
        ds_combo.grid(row=0, column=1, sticky="we", pady=(0, 4))
        ds_combo.bind("<<ComboboxSelected>>", lambda e: self.on_dataset_changed())

        tk.Label(main_frame, text="Starting letters:").grid(row=1, column=0, sticky="w")
        self.entry = tk.Entry(main_frame, textvariable=self.prefix_var, width=20)
        self.entry.grid(row=1, column=1, sticky="we", pady=(0, 6))
        self.entry.bind("<Control-Return>", self.on_ctrl_enter)
        
        status_label = tk.Label(main_frame, textvariable=self.status_var, fg="gray")
        status_label.grid(row=2, column=0, sticky="w", pady=(0, 6))
        
        self.roblox_status_label = tk.Label(main_frame, textvariable=self.roblox_status_var, fg="red")
        self.roblox_status_label.grid(row=2, column=1, sticky="e", pady=(0, 6))

        start_button = tk.Button(main_frame, text="Play round", command=self.on_play_round, bg="#e0e0e0")
        start_button.grid(row=3, column=0, sticky="we", pady=(0, 4), padx=(0,2))
        
        view_used_button = tk.Button(main_frame, text="View Used", command=self.show_used_words)
        view_used_button.grid(row=3, column=1, sticky="we", pady=(0, 4), padx=(2,0))

        speed_label = tk.Label(main_frame, text="Target Typing Speed (WPM):")
        speed_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        
        speed_scale = tk.Scale(
            main_frame, from_=20, to=300, orient="horizontal",
            variable=self.speed_var, showvalue=True, resolution=5
        )
        speed_scale.grid(row=5, column=0, columnspan=2, sticky="we", pady=(0, 4))

        typo_label = tk.Label(main_frame, text="Typo Chance (%):")
        typo_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(4, 0))
        
        typo_scale = tk.Scale(
            main_frame, from_=0.0, to=20.0, orient="horizontal",
            variable=self.typo_chance_var, showvalue=True, resolution=0.5
        )
        typo_scale.grid(row=7, column=0, columnspan=2, sticky="we", pady=(0, 4))

        mode_frame = tk.Frame(main_frame)
        mode_frame.grid(row=8, column=0, columnspan=2, sticky="we", pady=(8, 0))
        
        ttk.Label(mode_frame, text="Strategy:").pack(side='left')
        mode_combo = ttk.Combobox(
            mode_frame, textvariable=self.mode_var,
            values=["Random Words", "Short Words", "Long Words"],
            state='readonly', width=14
        )
        mode_combo.pack(side='left', padx=5)

        tk.Button(main_frame, text="Clear Cache", command=self.on_clear_cache).grid(row=9, column=0, sticky="we", pady=(12, 0), padx=(0,2))
        tk.Button(main_frame, text="Quit", command=self.root.destroy).grid(row=9, column=1, sticky="we", pady=(12, 0), padx=(2,0))

    def on_dataset_changed(self):
        self.status_var.set("Loading dataset...")
        selection = self.dataset_var.get()
        
        if "Combined" in selection:
            sets = ["web2", "gcide"]
        else:
            sets = [selection]
            
        threading.Thread(target=self._async_load, args=(sets,), daemon=True).start()

    def _async_load(self, sets: list[str]):
        self.db.load_dataset(sets)
        if self.db.is_loaded:
            self.root.after(0, lambda: self.status_var.set(f"Loaded {len(self.db.wordlist):,} words."))
        else:
            self.root.after(0, lambda: self.status_var.set(f"Failed to load: {self.db.error}"))

    def find_completion(self, prefix: str) -> str | None:
        candidates = self.db.get_candidates(prefix)
        
        valid = [w for w in candidates if w not in self.db.used_words and len(w) > len(prefix)]
        
        if not valid:
            return None

        mode = self.mode_var.get()
        if mode == "Short Words":
            chosen = min(valid, key=len)
        elif mode == "Long Words":
            chosen = max(valid, key=len)
        else:
            chosen = random.choice(valid)

        self.db.used_words.add(chosen)
        return chosen[len(prefix):]

    def _is_roblox_running(self) -> bool:
        try:
            startupinfo = None
            creationflags = 0
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW

            output = subprocess.check_output(
                ["tasklist"], text=True, startupinfo=startupinfo, creationflags=creationflags
            )
            return "RobloxPlayerBeta.exe" in output
        except Exception:
            return False

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
        if self._is_roblox_running():
            self.roblox_status_var.set("Roblox Found")
            self.roblox_status_label.config(fg="green")
        else:
            self.roblox_status_var.set("Roblox not Found")
            self.roblox_status_label.config(fg="red")

    def on_play_round(self) -> None:
        prefix = self.prefix_var.get().strip()
        if not prefix:
            messagebox.showwarning("Warning", "Enter starting letters.")
            return

        if not self.db.is_loaded:
            if self.db.error:
                messagebox.showerror("Error", f"Dictionary failed to load:\n{self.db.error}")
            else:
                messagebox.showwarning("Warning", "Dictionary still loading.")
            return

        completion = self.find_completion(prefix)
        if not completion:
            messagebox.showinfo("Result", "No unused words found for that prefix.")
            return

        self.root.withdraw()
        self.prefix_var.set("")
        self.refresh_roblox_status()
        
        if self._is_roblox_running():
            self._focus_roblox_window()

        threading.Thread(target=self._execute_typing, args=(completion,), daemon=True).start()

    def _execute_typing(self, completion: str) -> None:
        try:
            typo_prob = self.typo_chance_var.get() / 100.0
            typist = HumanTypist(target_wpm=self.speed_var.get(), typo_chance=typo_prob)
            typist.type_word(completion)
        finally:
            self.root.after(0, self.root.deiconify)
            if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
                self.root.after(0, self.update_used_words_list)

    def on_ctrl_enter(self, event):
        self.on_play_round()
        return "break"

    def on_clear_cache(self) -> None:
        self.db.used_words.clear()
        if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
            self.update_used_words_list()

    def show_used_words(self) -> None:
        if not hasattr(self, 'used_words_window') or not self.used_words_window.winfo_exists():
            self.used_words_window = tk.Toplevel(self.root)
            self.used_words_window.title(f"Used Words Session ({len(self.db.used_words)})")
            self.used_words_window.geometry("250x400")
            self.used_words_window.attributes('-topmost', True)
            
            frame = ttk.Frame(self.used_words_window)
            frame.pack(fill='both', expand=True, padx=5, pady=5)
            
            scrollbar = ttk.Scrollbar(frame)
            scrollbar.pack(side='right', fill='y')
            
            self.used_words_listbox = tk.Listbox(
                frame, yscrollcommand=scrollbar.set, font=('Consolas', 10)
            )
            self.used_words_listbox.pack(side='left', fill='both', expand=True)
            scrollbar.config(command=self.used_words_listbox.yview)
            
        self.update_used_words_list()
        self.used_words_window.lift()
        self.used_words_window.focus_force()
    
    def update_used_words_list(self) -> None:
        if hasattr(self, 'used_words_window') and self.used_words_window.winfo_exists():
            self.used_words_listbox.delete(0, tk.END)
            for word in sorted(self.db.used_words, key=str.lower):
                self.used_words_listbox.insert(tk.END, word)
            self.used_words_window.title(f"Used Words Session ({len(self.db.used_words)})")


def check_for_updates():
    try:
        response = requests.get("https://raw.githubusercontent.com/elDziad00/LastLetter/main/version.txt", timeout=3)
        if response.status_code == 200:
            latest = response.text.strip()
            if version.parse(latest) > version.parse(CURRENT_VERSION):
                print(f"Update available: {latest}")
    except Exception:
        pass

def main() -> None:
    root = tk.Tk()
    app = LastLetterApp(root)
    threading.Thread(target=check_for_updates, daemon=True).start()
    root.mainloop()

if __name__ == "__main__":
    main()