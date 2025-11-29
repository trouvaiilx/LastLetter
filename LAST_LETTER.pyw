import threading
import time
import random
import subprocess
import ctypes

import keyboard
import tkinter as tk
from tkinter import messagebox

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

        clear_cache_button = tk.Button(main_frame, text="Clear cache", command=self.on_clear_cache)
        clear_cache_button.grid(row=3, column=1, sticky="we", pady=(0, 4), padx=(4, 0))

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

        mode_label = tk.Label(main_frame, text="Word selection mode:")
        mode_label.grid(row=6, column=0, sticky="w")

        mode_dropdown = tk.OptionMenu(
            main_frame,
            self.mode_var,
            "Random Words",
            "Short Words",
            "Long Words",
        )
        mode_dropdown.grid(row=6, column=1, sticky="e")

        quit_button = tk.Button(main_frame, text="Quit", command=self.root.destroy)
        quit_button.grid(row=7, column=0, columnspan=2, sticky="we", pady=(0, 4))

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
            output = subprocess.check_output(["tasklist"], text=True)
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
        self.used_words.clear()

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


def main() -> None:
    root = tk.Tk()
    app = LastLetterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()