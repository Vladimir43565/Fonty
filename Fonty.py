import os
import shutil
import ctypes
import winreg
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Windows API Constants
HKEY_CURRENT_USER = winreg.HKEY_CURRENT_USER
USER_FONT_REG_PATH = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
WM_FONTCHANGE = 0x001D
HWND_BROADCAST = 0xFFFF

class FontyWord:
    def __init__(self):
        # Appearance Settings
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Fonty - Word")
        self.root.geometry("450x350")
        
        # UI Elements
        self.label = ctk.CTkLabel(self.root, text="Fonty - Word", font=("Arial", 24, "bold"))
        self.label.pack(pady=(20, 5))

        self.sublabel = ctk.CTkLabel(self.root, text="Drag & Drop Font Installer", font=("Arial", 12))
        self.sublabel.pack(pady=(0, 20))

        self.select_button = ctk.CTkButton(
            self.root, 
            text="Select .ttf or .otf Font", 
            height=50, 
            width=250,
            command=self.process_font
        )
        self.select_button.pack(pady=10)

        self.status_label = ctk.CTkLabel(self.root, text="Ready", text_color="gray")
        self.status_label.pack(pady=20)

    def restart_word(self):
        """Kills Word if open and restarts it to refresh font cache."""
        try:
            # Silently kill Word
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE", "/T"], capture_output=True)
            # Restart Word
            os.startfile("winword.exe")
            return True
        except Exception:
            # If Word wasn't running, just start it
            try:
                os.startfile("winword.exe")
                return True
            except:
                return False

    def install_font_local(self, font_path):
        """Installs font to Local AppData to avoid needing Admin rights."""
        font_full_name = os.path.basename(font_path)
        local_fonts_dir = os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'Fonts')
        
        if not os.path.exists(local_fonts_dir):
            os.makedirs(local_fonts_dir)

        dest_path = os.path.join(local_fonts_dir, font_full_name)

        # Fix for 'Permission Denied': Check if file exists
        if os.path.exists(dest_path):
            try:
                os.remove(dest_path)
                shutil.copy(font_path, dest_path)
            except OSError:
                # If we can't delete/overwrite, the font is currently in use by Windows/Word
                print("Font is currently locked by the system.")
        else:
            shutil.copy(font_path, dest_path)

        # Registry Integration
        clean_name = os.path.splitext(font_full_name)[0]
        reg_value_name = f"{clean_name} (TrueType)"
        
        with winreg.OpenKey(HKEY_CURRENT_USER, USER_FONT_REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, reg_value_name, 0, winreg.REG_SZ, font_full_name)

        # Notify Windows and GDI that font list changed
        ctypes.windll.user32.SendMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
        ctypes.windll.gdi32.AddFontResourceW(dest_path)
        return True

    def process_font(self):
        file_path = filedialog.askopenfilename(filetypes=[("Font Files", "*.ttf *.otf")])
        if file_path:
            try:
                if self.install_font_local(file_path):
                    self.status_label.configure(text="Font Installed Successfully!", text_color="#2ecc71")
                    
                    # Restart Logic
                    ask_restart = messagebox.askyesno("Success", "Font installed. Restart Word now to see your new font?")
                    if ask_restart:
                        self.restart_word()
                        self.status_label.configure(text="Word Restarted.")
            except Exception as e:
                self.status_label.configure(text="Installation Failed", text_color="#e74c3c")
                messagebox.showerror("Error", f"Failed to install: {e}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = FontyWord()
    app.run()