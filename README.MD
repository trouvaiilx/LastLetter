LastLetter Helpers
===================

External helper tools for the Roblox game "Last Letter".

They provide two modes:

- LAST_LETTER.pyw ("Last Letter" mode)
  - Classic last-letter style gameplay.
  - You type the starting letters, and it finds a word that begins with them.
  - It then types the rest of the word for you automatically.

- ONE_BY_ONE.pyw ("One By One" mode)
  - For games where you add one letter at a time to build a word.
  - You type the current fragment, and it suggests the next letter and a full word.

Safety
------

- These are external desktop tools written in Python.
- They do **not** inject into Roblox, its memory, or its process in any way.
- They only simulate normal keyboard input on your PC.
- You cannot get banned in Roblox for using these tools.

Getting the ready-made EXE (recommended)
----------------------------------------

For most users, simply download the latest `.exe` files from the GitHub **Releases** page and run them:

   LAST_LETTER.exe
   ONE_BY_ONE.exe

Running from source (advanced)
------------------------------

1. Create and activate a virtual environment (optional but recommended).
2. Install dependencies:

   pip install -r requirements.txt

3. Run one of the tools:

   python LAST_LETTER.pyw
   python ONE_BY_ONE.pyw

Building Windows executables (advanced)
---------------------------------------

Use PyInstaller to create your own standalone .exe files. Example commands:

   pyinstaller --onefile --windowed --icon=LastLetter.ico --add-data "LastLetter.ico;." --collect-data english_words LAST_LETTER.pyw

   pyinstaller --onefile --windowed --icon=ONEBYONE.ico --add-data "ONEBYONE.ico;." --collect-data english_words ONE_BY_ONE.pyw

The built executables will appear in the "dist" folder.
