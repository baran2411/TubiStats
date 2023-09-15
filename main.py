import tkinter as tk
from tkinter import ttk
import page_attendance
import page_goals
import page_assists
import page_minutes
import page_yellows
import page_penalties

def main():
    root = tk.Tk()
    root.title("Player Statistics Tracker")
    root.geometry("375x650")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    page_attendance.create_page(notebook)
    page_goals.create_page(notebook)
    page_assists.create_page(notebook)
    page_minutes.create_page(notebook)
    page_yellows.create_page(notebook)
    page_penalties.create_page(notebook)

    root.mainloop()

if __name__ == "__main__":
    main()