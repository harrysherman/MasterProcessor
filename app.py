import tkinter as tk
from tkinter import filedialog


def submit_files(file1, file2):
    print("File 1:", file1.get())
    print("File 2:", file2.get())


# Create the main window
root = tk.Tk()
root.title("Master Processor")
root.geometry("600x400")

# Set the background color to navy blue
root.configure(bg="#282F45")

# Variables to store selected file paths
master_file = tk.StringVar(value="None Selected")
lastest_export_file = tk.StringVar(value="None Selected")
test_running_toggle = tk.StringVar(value="Not Running")


def browse_master_file():
    file_path = filedialog.askopenfilename()
    master_file.set(file_path)
    print("selected master file", master_file.get())
    l1.configure(text=master_file.get())


def browse_export_file():
    file_path = filedialog.askopenfilename()
    lastest_export_file.set(file_path)
    print("selected file", lastest_export_file.get())
    l2.configure(text=lastest_export_file.get())


def both_files_selected():
    if master_file.get() != "None Selected" and lastest_export_file.get() != "None Selected":
        return "Cannot Run Code"
    else:
        return "Run Code"


# Entry widgets to display selected file paths
# Button to open the file browser
master_file_button = tk.Button(
    root, text="Select Master File", command=browse_master_file, bg="#B19557", width=20
)
master_file_button.pack()
l1 = tk.Label(root, text=(master_file.get()), bg="#282F45", fg="white")
l1.pack()

latest_export_button = tk.Button(
    root, text="Select Latest Export", command=browse_export_file, bg="#B19557", width=20
)
latest_export_button.pack()
l2 = tk.Label(root, text=(lastest_export_file.get()), bg="#282F45", fg="white")
l2.pack()

run_code_button = tk.Button(root, text=(both_files_selected()), bg="#B19557", width=20)
run_code_button.pack()


# Run the Tkinter event loop
root.mainloop()
