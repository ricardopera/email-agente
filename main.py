import sys
import os
from src.app import EmailApp
import tkinter as tk

def main():
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()