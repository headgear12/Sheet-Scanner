import pandas as pd
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

df = pd.read_excel('output.xlsx')

def process_comments(dataframe):
    comments = []
    for comment in dataframe['Comments']:
        if isinstance(comment, str):
            comments.extend(comment.split("\n")) 
    return pd.Series(comments).value_counts()  

root = tk.Tk()
root.title("Executor Comment Analysis")

def show_bar_graph(event):
    selected_executor = executor_combobox.get()

    filtered_df = df[df['Executor'] == selected_executor]

    comment_counts = process_comments(filtered_df).head(5)

    for widget in graph_frame.winfo_children():
        widget.destroy()

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(comment_counts.index, comment_counts.values, color='steelblue') 
    ax.set_xlabel("Frequency")
    ax.set_ylabel("Comments")
    ax.set_title(f"Top 5 Comments for Executor: {selected_executor}")
    ax.invert_yaxis()

    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

executor_combobox = ttk.Combobox(root, values=df['Executor'].unique().tolist())
executor_combobox.set("Select Executor")
executor_combobox.bind("<<ComboboxSelected>>", show_bar_graph)

executor_combobox.pack(pady=20)

graph_frame = tk.Frame(root)
graph_frame.pack(pady=20)

root.mainloop()
