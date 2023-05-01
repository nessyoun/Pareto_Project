import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter


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
        df['cumperc'] = df[nom_col].cumsum()/df[nom_col].sum()*100
        
        #define aesthetics for plot
        color1 = 'steelblue'
        color2 = 'red'
        line_size = 4
        
        #create basic bar plot
        fig, ax = plt.subplots()
        ax.bar(df.index, df[nom_col], color=color1)
        
        #add cumulative percentage line to plot
        ax2 = ax.twinx()
        ax2.plot(df.index, df['cumperc'], color=color2, marker="D", ms=line_size)
        ax2.yaxis.set_major_formatter(PercentFormatter())
        
        #specify axis colors
        ax.tick_params(axis='y', colors=color1)
        ax2.tick_params(axis='y', colors=color2)
        droit_quatre_ving=ax2.axhline(y=80, color='r', linestyle='--')
        #display Pareto chart
        plt.show()
        
#Pareto("data.xlsx", "Feuil1",'cout_par_occurrence' ).create_pareto()
