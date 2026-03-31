Fonty - Word
The easiest way to add custom fonts to Microsoft Word.

Fonty - Word is a lightweight Python utility designed to bypass the manual, multi-step process of installing fonts on Windows. It allows you to select a font and automatically handles the directory placement, registry registration, and system notification required for Word to recognize it.

Key Features
No Admin Required: Installs fonts to the local user directory, meaning you don't need administrator privileges or UAC popups to use it.

One-Click Word Restart: Offers to automatically close and relaunch Microsoft Word so your new font appears in the dropdown menu immediately.

Modern Interface: Built with a clean, dark-mode GUI using CustomTkinter.

System Integration: Uses Windows API calls to notify the OS of font changes without requiring a full system reboot.

How to Use
Run the App: Open Fonty-Word.exe (or run the script via Python).

Select Font: Click the Select .ttf or .otf Font button.

Install: Choose your font file (e.g., Minecraft.ttf).

Restart Word: When prompted, click Yes to let Fonty restart Microsoft Word for you.

Enjoy: Open your font list in Word and look for your new font.

Installation for Developers
If you want to run the source code directly, follow these steps:

1. Prerequisites
Ensure you have Python 3.x installed. You will need the customtkinter library:

Bash
pip install customtkinter
2. Running the Script
Download fonty.py and run:

Bash
python fonty.py
3. Creating a Standalone .exe
To turn this script into a single executable file that you can share:

Bash
pip install pyinstaller
pyinstaller --noconsole --onefile fonty.py
The finished app will be located in the dist folder.

Troubleshooting
Permission Denied: Ensure Microsoft Word is closed before installing a font. If Word is using the font, Windows will lock the file and prevent Fonty from updating it.

Font Not Showing: Some fonts require a full Word restart. Use the built-in "Restart Word" prompt to ensure the font cache is refreshed.

Supported Formats: Currently supports .ttf (TrueType) and .otf (OpenType) files.

License
This project is for personal use. Please ensure you have the proper legal license for any custom fonts you install.