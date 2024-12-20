import tkinter as tk
import os

os.environ['TKINTER_VERBOSE'] = '1'  # Enable debug mode

print("Creating window...")
root = tk.Tk()
print("Window created")
root.title("Test")
print("Set title")
root.geometry("200x200")
print("Set geometry")
label = tk.Label(root, text="Test Window")
print("Created label")
label.pack()
print("Packed label")
print("Starting mainloop...")
root.mainloop()
print("Mainloop ended")