import threading

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


class OneByOneHelperApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("One By One Helper")
        self.root.resizable(False, False)

        try:
            self.root.iconbitmap(resource_path("ONEBYONE.ico"))
        except Exception as e:
            print("Icon load error:", e)


        self.wordlist: list[str] = []
        self.wordlist_loaded = False
        self.wordlist_error = None

        self.fragment_var = tk.StringVar()
        self.result_var = tk.StringVar(value="Waiting for word list to load...")

        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)

        label = tk.Label(main_frame, text="Current fragment:")
        label.grid(row=0, column=0, sticky="w")

        entry = tk.Entry(main_frame, textvariable=self.fragment_var, width=20)
        entry.grid(row=1, column=0, columnspan=2, sticky="we", pady=(2, 6))
        entry.focus_set()
        entry.bind("<Control-Return>", self.on_ctrl_enter)

        suggest_button = tk.Button(main_frame, text="Suggest move", command=self.on_suggest)
        suggest_button.grid(row=2, column=0, sticky="we", pady=(0, 4))

        clear_button = tk.Button(main_frame, text="Clear", command=self.on_clear)
        clear_button.grid(row=2, column=1, sticky="we", pady=(0, 4), padx=(4, 0))

        result_label_title = tk.Label(main_frame, text="Suggestion:")
        result_label_title.grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        result_label = tk.Label(main_frame, textvariable=self.result_var, justify="left", anchor="w")
        result_label.grid(row=4, column=0, columnspan=2, sticky="we")

        quit_button = tk.Button(main_frame, text="Quit", command=self.root.destroy)
        quit_button.grid(row=5, column=0, columnspan=2, sticky="we", pady=(6, 0))

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        threading.Thread(target=self.load_wordlist, daemon=True).start()

    def load_wordlist(self) -> None:
        try:
            words_set = get_english_words_set(["web2"], lower=False)
            self.wordlist = sorted(w for w in words_set if w.isalpha())
            self.wordlist_loaded = True
            self.wordlist_error = None
            self.root.after(0, lambda: self.result_var.set("Word list loaded. Enter a fragment."))
        except Exception as exc:
            self.wordlist_loaded = False
            self.wordlist_error = exc
            self.root.after(0, lambda: self.result_var.set("Failed to load word list."))

    def find_best_completion(self, fragment: str) -> tuple[str | None, str | None]:
        if not self.wordlist_loaded or not self.wordlist:
            return None, None

        frag_lower = fragment.lower()

        best_word: str | None = None
        for word in self.wordlist:
            if not word.lower().startswith(frag_lower):
                continue
            if len(word) <= len(fragment):
                continue
            if best_word is None or len(word) < len(best_word):
                best_word = word
        if best_word is None:
            return None, None

        next_letter = best_word[len(fragment)]
        return next_letter, best_word

    def on_suggest(self) -> None:
        fragment = self.fragment_var.get().strip()
        if not fragment:
            messagebox.showwarning("One By One Helper", "Please enter the current fragment.")
            return

        if not self.wordlist_loaded:
            messagebox.showwarning(
                "One By One Helper", "Word list is still loading or failed. Try again in a moment.",
            )
            return

        next_letter, full_word = self.find_best_completion(fragment)
        if next_letter is None or full_word is None:
            self.result_var.set("No valid completion found for that fragment.")
            return

        self.result_var.set(
            f"Next letter: {next_letter}\nFull word: {full_word}",
        )

    def on_ctrl_enter(self, event):
        self.on_suggest()
        return "break"

    def on_clear(self) -> None:
        self.fragment_var.set("")
        self.result_var.set("Enter a fragment.")


def main() -> None:
    root = tk.Tk()
    app = OneByOneHelperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()