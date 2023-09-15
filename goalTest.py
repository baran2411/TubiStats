import tkinter as tk
from tkinter import messagebox
import openpyxl

# Define a dictionary to map player names to player IDs
player_id_dict = {
    "Baran": "1",
    "Cem": "2",
    "Emir": "3",
    "Ersan": "4",
    "Ivo": "5",
    "Jochem": "6",
    "Jordy": "7",
    "Justin": "8",
    "Kaan": "9",
    "Kevin": "10",
    "Koray": "11",
    "Lars": "12",
    "Luc": "13",
    "Oseb": "14",
    "Rick": "15",
    "Stefan": "16",
    "Sven": "17",
    "Swen": "18"
}

# List of player names
players = list(player_id_dict.keys())

player_vars = []

# Function to record attendance
def record_attendance():
    # Get the selected players (those present) and their IDs
    present_players = [player_id_dict[players[i]] for i, var in enumerate(player_vars) if var.get()]

    # Get the training ID
    training_id = training_id_entry.get()

    # Write the data to the Excel file
    excel_file = "Goals.xlsx"
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    # Append each player's ID as a separate row
    for player_id in present_players:
        sheet.append([training_id, player_id])

    # Save the Excel file
    workbook.save(excel_file)

    # Reset the form
    training_id_entry.delete(0, tk.END)
    for var in player_vars:
        var.set(False)

    # Show a success message
    messagebox.showinfo("Success", "Attendance recorded successfully!")

# Create the main application window
root = tk.Tk()
root.title("Training Attendance Form")

# Label and entry for Training ID
training_id_label = tk.Label(root, text="Training ID:")
training_id_label.pack()
training_id_entry = tk.Entry(root)
training_id_entry.pack()

# Checkboxes for selecting players (18 players)
for player in players:
    var = tk.BooleanVar()
    player_vars.append(var)
    checkbox = tk.Checkbutton(root, text=player, variable=var)
    checkbox.pack()

# Button to record attendance
record_button = tk.Button(root, text="Record Attendance", command=record_attendance)
record_button.pack()

# Start the GUI application
root.mainloop()
