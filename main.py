import numpy as np

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from openpyxl import load_workbook

import tkinter as tk
from tkinter import *
from tkinter import filedialog

matplotlib.use('TkAgg')

# process the excel given the filepath
def processExcel(path):
    wb = load_workbook(path)
    ws = wb['Sheet1']
    all_rows = list(ws.rows)

    data = []
    for row in all_rows[1:]:
        value = row[1].value
        if (value is not None):
            data.append(value)

    n, bins, patches = plt.hist(data, bins=40, alpha=0.5, color='r')

    plt.xlabel('Edad', fontweight ="bold")
    plt.ylabel('Frecuencia', fontweight ="bold")
    plt.title('Edad de los perrunos', fontweight ="bold")

    fig = plt.figure(1)
    canvas = FigureCanvasTkAgg(fig, master=window)
    plot_widget = canvas.get_tk_widget()
    plot_widget.grid(column=0, row=2)

    data2 = []
    for row in all_rows[1:]:
        value = row[2].value
        if (value is not None):
            data2.append(value)

    plt.figure(figsize=(12,6))
    n, bins, patches = plt.hist(data2, bins=40, alpha=0.5, color='r')

    plt.xlabel('Raza', fontweight ="bold")
    plt.ylabel('Frecuencia', fontweight ="bold")
    plt.title('Raza de los perrunos', fontweight ="bold")
    plt.xticks(rotation='vertical')
    plt.subplots_adjust(bottom=0.35)

    fig = plt.figure(2)
    canvas = FigureCanvasTkAgg(fig, master=window)
    plot_widget = canvas.get_tk_widget()
    plot_widget.grid(column=1, row=2)


# method to brose the file in OS
def browseFiles():
    filepath = filedialog.askopenfilename(
        initialdir = "C:/Users/Jorge/Desktop/PerritosApp", 
        title = "Select a File", 
        filetypes = (("Text files", "*.xlsx*"), ("all files", "*.*")))

    if not filepath:
        messagebox.showerror("Error", "Selecciona un archivo")
        return

    lblFileRoute.configure(text="Path of the file: " + filepath)
    processExcel(filepath)

def center_window(width=300, height=200):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry('%dx%d+%d+%d' % (width, height, x, y))

# Create the window
window = tk.Tk()
window.title('Perrunos App')
window.geometry("1800x800")
window.config(background = "white")

center_window(1800, 800)

# Create widgets of window
lblFileExplorer = Label(window, text="Perrunos App", width=50, height=3, fg="black")
lblFileRoute = Label(window, text="Ruta del archivo", width=80, height=2, fg="red")
button_explore = Button(window, text="Browse Files", command=browseFiles)
button_exit = Button(window, text="Exit", command=exit)
  
# Locate widgets on window
lblFileExplorer.grid(column=0, row=0)
button_explore.grid(column=0, row=1)
lblFileRoute.grid(column=0, row=3)
button_exit.grid(column=0, row=4)

window.mainloop()