import tkinter as tk
from tkinter import messagebox
import csv
from datetime import datetime
import os
import locale

locale.setlocale(locale.LC_TIME, 'nl_NL.UTF-8')

def bezoekers_registratie():
    bezoekers = entry.get()
    
    if not bezoekers.isdigit():
        messagebox.showerror("Invalid input", "Voer een getal in.")
        return

    today = datetime.now()
    day_of_week = today.strftime('%A')
    file_exists = os.path.isfile('bezoekers_registratie.csv')

    with open('bezoekers_registratie.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(['Datum', 'Dag', 'Aantal Bezoekers'])
        writer.writerow([today.strftime('%d-%m-%Y'), day_of_week, bezoekers])

    messagebox.showinfo("Success", f"Er zijn {bezoekers} bezoekers geregistreerd op {day_of_week}.")
    entry.delete(0, tk.END)

# Create the main window
root = tk.Tk()
root.title("Bezoekers registratie")

# Create and place the widgets
label = tk.Label(root, text="Voer het aantal bezoekers in:")
label.pack(pady=10)

entry = tk.Entry(root)
entry.pack(pady=5)

button = tk.Button(root, text="Bezoekers registratie", command=bezoekers_registratie)
button.pack(pady=20)

# Start the GUI event loop
root.mainloop()