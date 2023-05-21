import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import PercentFormatter
from matplotlib.patches import Rectangle

class Pareto:
    def __init__(self, nomfic, nomfeuil, column):
        self.nomfic = nomfic
        self.nomfeuil = nomfeuil
        self.column = column

    def create_pareto(self):
        nom_col = None
        df = pd.read_excel(self.nomfic, sheet_name=self.nomfeuil, index_col=0)
        page = df.columns.get_loc(self.column)
        nom_col = df.columns[page]
        df = df.sort_values(by=nom_col, ascending=False)
        df['cumperc'] = df[nom_col].cumsum() / df[nom_col].sum() * 100

        # define aesthetics for plot
        color1 = 'steelblue'
        color2 = 'red'
        line_size = 4

        # create basic bar plot
        fig = plt.Figure(figsize=(6, 4), dpi=100)  # Modify the figsize as per your requirements
        ax = fig.add_subplot(111)
        bars = ax.bar(df.index, df[nom_col], color=color1)

        # set x-axis tick positions
        x_ticks = np.arange(len(df.index))

        # set x-axis tick labels with rotation
        ax.set_xticks(x_ticks)
        ax.set_xticklabels(df.index, rotation=90)

        # add cumulative percentage line to plot
        ax2 = ax.twinx()
        ax2.plot(df.index, df['cumperc'], color=color2, marker="D", ms=line_size)
        ax2.yaxis.set_major_formatter(PercentFormatter())

        # specify axis colors
        ax.tick_params(axis='y', colors=color1)
        ax2.tick_params(axis='y', colors=color2)
        droit_quatre_ving = ax2.axhline(y=80, color='r', linestyle='--')

        # find intersection point
        intersection_x = np.interp(80, df['cumperc'], x_ticks)

        # annotate intersection point
        ax.annotate(f'Intersection: {df.index[int(intersection_x)]}', xy=(intersection_x, 80),
                    xytext=(10, 10), textcoords='offset points',
                    arrowprops=dict(arrowstyle='->'))

        # add colored rectangles
        left_rectangle = Rectangle((0, 0), intersection_x, ax.get_ylim()[1],
                                   facecolor='red', alpha=0.3)
        right_rectangle = Rectangle((intersection_x, 0), len(df.index) - intersection_x,
                                    ax.get_ylim()[1], facecolor='green', alpha=0.3)

        ax.add_patch(left_rectangle)
        ax.add_patch(right_rectangle)

        # create Tkinter window
        root = tk.Tk()
        root.title("Pareto Chart")

        # calculate window size based on the size of the figure
        window_width = fig.get_size_inches()[0] * fig.dpi + 50
        window_height = fig.get_size_inches()[1] * fig.dpi + 50

        # set the size of the Tkinter window
        root.geometry(f"{int(window_width)}x{int(window_height)}")

        # create FigureCanvasTkAgg instance
        canvas = FigureCanvasTkAgg(fig, master=root)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # create a horizontal scrollbar
        hscrollbar = tk.Scrollbar(root, orient=tk.HORIZONTAL)
        hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # configure the scrollbar to scroll the canvas horizontally
        hscrollbar.config(command=canvas.get_tk_widget().xview)
        canvas.get_tk_widget().config(xscrollcommand=hscrollbar.set)

        # create a new window for the list of bar names
        bar_names_window = tk.Toplevel(root)
        bar_names_window.title("Bar Names")

        # create a text widget to display the bar names
        bar_names_text = tk.Text(bar_names_window)
        bar_names_text.pack(fill=tk.BOTH, expand=True)

        # get the list of bar names to display
        bar_names = df.index[:int(intersection_x)][:]  # Reverse the order

        # insert the bar names into the text widget
        counteur=0
        bar_names_text.insert(tk.END, "Les causes du probleme sont trier par le plus dangereux au moins:\n\n\n")
        for name in bar_names:
            counteur+=1
            bar_names_text.insert(tk.END, str(counteur)+") "+name + "\n")

        # start Tkinter event loop
        tk.mainloop()

#Pareto("data.xlsx", "Feuil1", 'cout_par_occurrence').create_pareto()
