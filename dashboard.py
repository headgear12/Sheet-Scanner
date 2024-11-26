import pandas as pd
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from datetime import datetime

df = pd.read_excel('output.xlsx')

df['Executed Date'] = pd.to_datetime(df['Executed Date'], format='%b %d, %Y', errors='coerce')
df['Reviewed Date'] = pd.to_datetime(df['Reviewed Date'], format='%b %d, %Y', errors='coerce')
df['Execution Year'] = df['Executed Date'].dt.year
df['Execution Month'] = df['Executed Date'].dt.month_name()
df['Review Year'] = df['Reviewed Date'].dt.year
df['Review Time'] = (df['Reviewed Date'] - df['Executed Date']).dt.days

df = df.dropna(subset=['Execution Year', 'Review Year'])

execution_counts = df.groupby('Execution Year').size().astype(int).sort_index()
review_counts = df.groupby('Review Year').size().astype(int).sort_index()

current_year = datetime.now().year
df_current_year = df[df['Execution Year'] == current_year]

monthly_execution_counts = df_current_year.groupby('Execution Month').size().reindex([
    'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'
], fill_value=0).astype(int)

executor_counts = df['Executor'].value_counts().astype(int)
reviewer_counts = df['Reviewer'].value_counts().astype(int)

def plot_yearly_trends(frame):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(execution_counts.index, execution_counts.values, label='Execution Dates', color='b', marker='o', linestyle='-', linewidth=2)
    ax.plot(review_counts.index, review_counts.values, label='Review Dates', color='g', marker='o', linestyle='-', linewidth=2)
    ax.set_xlabel('Year')
    ax.set_ylabel('Count')
    ax.set_title('Yearly Trends of Execution and Review')
    ax.legend()
    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

def plot_monthly_execution(frame):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(monthly_execution_counts.index, monthly_execution_counts.values, color='b')
    ax.set_xlabel('Month')
    ax.set_ylabel('Count')
    ax.set_title(f'Monthly Execution Counts ({current_year})')
    plt.xticks(rotation=45)
    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

def plot_executor_pie(frame):
    fig, ax = plt.subplots(figsize=(5, 4))
    ax.pie(executor_counts, labels=executor_counts.index, autopct=lambda p: f'{int(p * sum(executor_counts) / 100)}', startangle=90, colors=plt.cm.Paired.colors)
    ax.set_title('Execution by Executor')
    ax.axis('equal')

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

def plot_reviewer_pie(frame):
    fig, ax = plt.subplots(figsize=(5, 4))
    ax.pie(reviewer_counts, labels=reviewer_counts.index, autopct=lambda p: f'{int(p * sum(reviewer_counts) / 100)}', startangle=90, colors=plt.cm.Paired.colors)
    ax.set_title('Review by Reviewer')
    ax.axis('equal')

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

root = tk.Tk()
root.title("Design Analysis Dashboard")
root.geometry("1200x800")

yearly_trends_frame = tk.Frame(root)
monthly_execution_frame = tk.Frame(root)
executor_pie_frame = tk.Frame(root)
reviewer_pie_frame = tk.Frame(root)

yearly_trends_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
monthly_execution_frame.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
executor_pie_frame.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')
reviewer_pie_frame.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')

root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)

plot_yearly_trends(yearly_trends_frame)
plot_monthly_execution(monthly_execution_frame)
plot_executor_pie(executor_pie_frame)
plot_reviewer_pie(reviewer_pie_frame)

root.mainloop()
