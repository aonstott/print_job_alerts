import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import PhotoImage
import subprocess
import os

save_files = "no"

def on_checkbox_toggle():
    global save_files
    if checkbox_var.get():
        save_files = "yes"
    else:
        save_files = "no"

def run_script(send_emails):
    filename = file_entry.get()
    if filename:
        try:
            abs_path = os.path.abspath(filename)
            main_path = os.path.join(os.path.dirname(__file__), "main.py")
            subprocess.run(["python3", main_path, filename, send_emails, save_files], check=True)
        except subprocess.CalledProcessError as e:
            result_text.set("Error: {}".format(e))
    else:
        result_text.set("Please select a file")

# Create the main window
root = tk.Tk()
root.title("PM Email Alerts")

icon = PhotoImage(file="appicon.png")
root.iconphoto(True, icon)
# Create a label and entry for the filename
file_label = tk.Label(root, text="Filename:")
file_label.grid(row=0, column=0, padx=5, pady=5)

file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)

# Create a button to select the file
file_button = tk.Button(root, text="Select File", command=lambda: file_entry.insert(tk.END, filedialog.askopenfilename()))
file_button.grid(row=0, column=2, padx=5, pady=5)

# Create a button to run the script
run_button = tk.Button(root, text="Download Files", command=lambda: run_script("test"))
run_button.grid(row=1, column=1, padx=5, pady=5)

# Create a button to send the emails
send_button = tk.Button(root, text="Send Emails", command=lambda: run_script("send_emails"))
send_button.grid(row=2, column=1, padx=5, pady=5)

# Create a label to display the result
result_text = tk.StringVar()
result_label = tk.Label(root, textvariable=result_text, fg="blue")
result_label.grid(row=3, column=1, padx=5, pady=5)

checkbox_var = tk.BooleanVar()
checkbox = ttk.Checkbutton(root, text="Save Files", variable=checkbox_var, command=on_checkbox_toggle)
checkbox.grid(row=0, column=0, padx=10, pady=10)

# Start the GUI main loop
root.mainloop()