import tkinter as tk
from tkinter import filedialog
import subprocess
import os

def run_script():
    filename = file_entry.get()
    if filename:
        try:
            abs_path = os.path.abspath(filename)
            main_path = os.path.join(os.path.dirname(__file__), "main.py")
            subprocess.run(["python3", main_path, filename], check=True)
        except subprocess.CalledProcessError as e:
            result_text.set("Error: {}".format(e))
    else:
        result_text.set("Please select a file")

# Create the main window
root = tk.Tk()
root.title("Email Alerts for Unfinished Jobs")

# Create a label and entry for the filename
file_label = tk.Label(root, text="Filename:")
file_label.grid(row=0, column=0, padx=5, pady=5)

file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)

# Create a button to select the file
file_button = tk.Button(root, text="Select File", command=lambda: file_entry.insert(tk.END, filedialog.askopenfilename()))
file_button.grid(row=0, column=2, padx=5, pady=5)

# Create a button to run the script
run_button = tk.Button(root, text="Run Script", command=run_script)
run_button.grid(row=1, column=1, padx=5, pady=5)

# Create a label to display the result
result_text = tk.StringVar()
result_label = tk.Label(root, textvariable=result_text, fg="blue")
result_label.grid(row=2, column=1, padx=5, pady=5)

# Start the GUI main loop
root.mainloop()