import tkinter as tk


def on_button_click():
    label.config(text="Hello, Tkinter!")


# Create the main window
app = tk.Tk()
app.title("Minimal Tkinter App")

# Create a label
label = tk.Label(app, text="Welcome to Tkinter!")
label.pack(pady=10)

# Create a button
button = tk.Button(app, text="Click me!", command=on_button_click)
button.pack(pady=10)

# Run the Tkinter event loop
app.mainloop()
