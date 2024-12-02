import tkinter as tk
from tkinter import messagebox
import subprocess

def run_scan_pdf():
    screen.withdraw()
    try:
        process = subprocess.Popen(["python", "Project.py"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        process.communicate()
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to run {e.cmd}.\n{e.output}")
    finally:
        screen.deiconify()

def show_dashboard():
    screen.withdraw()
    try:
        process = subprocess.Popen(["python", "dashboard.py"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        process.communicate()
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to run {e.cmd}.\n{e.output}")
    finally:
        screen.deiconify()

def view_employee_performance():
    screen.withdraw()
    try:
        process = subprocess.Popen(["python", "frequency.py"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        process.communicate()
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to run {e.cmd}.\n{e.output}")
    finally:
        screen.deiconify()

def create_gui():
    global screen
    screen = tk.Tk()
    screen.geometry("400x600+600+50")
    screen.title("Select Option")
    screen.configure(background="#0a3542")

    tk.Label(text="Pic2Sheet", font=("Arial", 30), bg="#0a3542", fg="white").pack(pady=50)
    tk.Button(text="Import PDF", font=("Arial", 20), height=2, width=20, bg="#5c909c", fg="white", bd=0, command=run_scan_pdf).pack(pady=20)
    tk.Button(text="Employee Insights", font=("Arial", 20), height=2, width=20, bg="#5c909c", fg="white", bd=0, command=view_employee_performance).pack(pady=20)
    tk.Button(text="Show Dashboard", font=("Arial", 20), height=2, width=20, bg="#5c909c", fg="white", bd=0, command=show_dashboard).pack(pady=20)
    screen.mainloop()

create_gui()
