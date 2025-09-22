import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.geometry("300x200")
label = ttk.Label(root, text="Test Label")
label.pack(pady=20)
button = ttk.Button(root, text="Test Button")
button.pack()
root.mainloop()