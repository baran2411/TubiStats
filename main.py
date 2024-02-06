import tkinter as tk
from tkinter import ttk
from pages import training_attendance
from pages import goal
from pages import assist
from pages import minutes
from pages import yellow_card
from pages import penalty

def main():
    root = tk.Tk()
    root.title("Tubi Statsheet")
    root.geometry("375x660")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    training_attendance.create_page(notebook)
    minutes.create_page(notebook)
    goal.create_page(notebook)
    assist.create_page(notebook)
    yellow_card.create_page(notebook)
    penalty.create_page(notebook)

    root.mainloop()

if __name__ == "__main__":
    main()