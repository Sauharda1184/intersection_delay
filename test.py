import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("Test Tkinter")
root.geometry("400x200")

label = ttk.Label(root, text="Hello, Tkinter!", font=("Helvetica", 16))
label.pack(pady=50)

root.mainloop()
