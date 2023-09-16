import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from openpyxl import load_workbook
import players  # Import the player data module

# Function to record minutes
def record_minutes():
    try:
        # Specify the path to the original Excel file for minutes
        minutes_excel_file = "F:\OneDrive\TubiStats\Minute.xlsx"

        # Load the existing workbook
        workbook = load_workbook(minutes_excel_file)

        # Get the active sheet (you can modify this to select a specific sheet)
        sheet = workbook.active

        # Determine the next match ID based on the last value in the "matchId" column
        last_match_id = sheet.cell(row=sheet.max_row, column=3).value
        if last_match_id is None:
            match_id = 1
        else:
            match_id = last_match_id + 1

        # Get the player names and the corresponding minutes
        minutes_data = []
        for i, player in enumerate(players.players):
            minutes = minutes_vars[i].get()
            if minutes:
                player_id = int(players.player_id_dict[player])  # Wrap in int() to ensure it's an integer
                minutes = int(minutes)  # Wrap in int() to ensure it's an integer
                minutes_data.append([match_id, player_id, minutes])

        # Calculate the starting index based on existing data
        existing_data = [cell.value for cell in sheet["A"]]
        if all(value is None for value in existing_data[1:]) or (len(existing_data) == 1 and existing_data[0] is None):
            start_index = 1
        else:
            start_index = max([int(value) for value in existing_data[1:] if value is not None], default=0) + 1

        # Append each player's data as a separate row with the specified index and the determined match ID
        for i, (match_id, player_id, minutes) in enumerate(minutes_data, start=start_index):
            sheet.append([i, player_id, match_id, minutes])

        # Save the Excel file with the updated data
        workbook.save(minutes_excel_file)

        # Reset the input fields
        for minutes_var in minutes_vars:
            minutes_var.set("")

        # Show a success message
        messagebox.showinfo("Success", "Minutes recorded successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def create_page(notebook):
    minutes_frame = ttk.Frame(notebook)
    notebook.add(minutes_frame, text="Minutes")

    global minutes_vars
    minutes_vars = []

    # Create labels and input fields for goal recording
    bold_font = ("", 10, "bold")  # Define a bold font

    # Create labels and input fields for minute recording
    for player in players.players:
        label = ttk.Label(minutes_frame, text=player, font=bold_font)
        label.grid(column=0, row=players.players.index(player), padx=10, pady=5, sticky=tk.W)

        var = tk.StringVar()
        minutes_vars.append(var)

        minutes_entry = ttk.Entry(minutes_frame, textvariable=var)
        minutes_entry.grid(column=1, row=players.players.index(player), padx=10, pady=5, sticky=tk.E)

    # Button to record minutes - Minutes Page
    minutes_button = ttk.Button(minutes_frame, text="Confirm", command=record_minutes)
    minutes_button.grid(column=0, row=len(players.players), columnspan=2, padx=10, pady=10)
    minutes_button.configure(width=20)  # Adjust the button width for better appearance

    # Center the button within the frame
    minutes_frame.columnconfigure(0, weight=1)