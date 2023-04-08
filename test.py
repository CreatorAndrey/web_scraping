import tkinter as tk
import tkinter.ttk as ttk
import time

# Create the master object
master = tk.Tk()

# Create a progressbar widget
progress_bar = ttk.Progressbar(master, orient="horizontal",
                               mode="determinate", maximum=100, value=0)

# And a label for it
label_1 = tk.Label(master, text="Progress Bar")

# Use the grid manager
label_1.grid(row=0, column=0)
progress_bar.grid(row=0, column=1)

# Necessary, as the master object needs to draw the progressbar widget
# Otherwise, it will not be visible on the screen
master.update()

progress_bar['value'] = 0
master.update()

while progress_bar['value'] < 100:
    progress_bar['value'] += 10
    # Keep updating the master object to redraw the progress bar
    master.update()
    time.sleep(0.5)

# The application mainloop
tk.mainloop()