import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
import os

class GraphApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Gráficos desde Excel")

        # Variables
        self.filepath = tk.StringVar()
        self.savepath = tk.StringVar()
        self.selected_x = tk.StringVar()
        self.selected_y = tk.StringVar(value=[])
        self.chart_type = tk.StringVar()
        self.legend = tk.BooleanVar(value=True)
        self.individual_graphs = tk.BooleanVar(value=False)  # Checkbox for individual graphs

        # UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Select Excel File
        btn_open = tk.Button(self.root, text="Seleccionar archivo Excel", command=self.load_excel)
        btn_open.pack(pady=10)

        # Dropdown for X axis
        lbl_x = tk.Label(self.root, text="Selecciona la columna para el eje X:")
        lbl_x.pack(pady=5)
        self.ddl_x = ttk.Combobox(self.root, textvariable=self.selected_x)
        self.ddl_x.pack(pady=5)

        # Listbox for Y axis columns
        lbl_y = tk.Label(self.root, text="Selecciona las columnas para el eje Y:")
        lbl_y.pack(pady=5)
        self.listbox_y = tk.Listbox(self.root, selectmode=tk.MULTIPLE)
        self.listbox_y.pack(pady=5)

        # Checkbox for individual graphs
        cb_individual = tk.Checkbutton(self.root, text="Gráficos individuales", variable=self.individual_graphs)
        cb_individual.pack(pady=5)

        # Chart type
        lbl_chart_type = tk.Label(self.root, text="Tipo de Gráfico:")
        lbl_chart_type.pack(pady=5)
        self.ddl_chart_type = ttk.Combobox(self.root, values=["line", "bar", "scatter", "scatter-line"], textvariable=self.chart_type)
        self.ddl_chart_type.pack(pady=5)

        # Legend Checkbox
        cb_legend = tk.Checkbutton(self.root, text="Incluir leyenda", variable=self.legend)
        cb_legend.pack(pady=5)

        # Save Path
        btn_save_path = tk.Button(self.root, text="Seleccionar carpeta para guardar", command=self.set_save_path)
        btn_save_path.pack(pady=10)

        # Generate Graph
        btn_generate = tk.Button(self.root, text="Generar gráfico", command=self.generate_graph)
        btn_generate.pack(pady=20)

    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not filepath:
            return
        self.filepath.set(filepath)

        # Load Excel using pandas and populate dropdowns
        self.df = pd.read_excel(filepath)
        columns = self.df.columns.tolist()

        # Set the values for dropdown and listbox
        self.ddl_x["values"] = columns
        for col in columns:
            self.listbox_y.insert(tk.END, col)

    def set_save_path(self):
        folder_selected = filedialog.askdirectory()
        self.savepath.set(folder_selected)

    def generate_graph(self):
        if not self.filepath.get():
            messagebox.showerror("Error", "Por favor, selecciona un archivo Excel.")
            return
        if not self.savepath.get():
            messagebox.showerror("Error", "Por favor, selecciona una carpeta para guardar.")
            return
        if not self.selected_x.get():
            messagebox.showerror("Error", "Por favor, selecciona una columna para el eje X.")
            return
        if not self.listbox_y.curselection():
            messagebox.showerror("Error", "Por favor, selecciona al menos una columna para el eje Y.")
            return
        if not self.chart_type.get():
            messagebox.showerror("Error", "Por favor, selecciona un tipo de gráfico.")
            return

        x = self.df[self.selected_x.get()]
        y_cols = [self.listbox_y.get(i) for i in self.listbox_y.curselection()]

        if self.individual_graphs.get():
            for y_col in y_cols:
                plt.figure(figsize=(10, 6))
                if self.chart_type.get() == "scatter-line":
                    plt.scatter(x, self.df[y_col], label=y_col)
                    plt.plot(x, self.df[y_col])
                elif self.chart_type.get() == "line":
                    plt.plot(x, self.df[y_col], label=y_col)
                elif self.chart_type.get() == "scatter":
                    plt.scatter(x, self.df[y_col], label=y_col)
                plt.xlabel(self.selected_x.get())
                plt.ylabel(y_col)
                if self.legend.get():
                    plt.legend()
                filename = os.path.join(self.savepath.get(), f"{y_col}.svg")
                plt.savefig(filename, format='svg',bbox_inches='tight')
                plt.close()
        else:
            plt.figure(figsize=(10, 6))
            for y_col in y_cols:
                if self.chart_type.get() == "scatter-line":
                    plt.scatter(x, self.df[y_col], label=y_col)
                    plt.plot(x, self.df[y_col])
                elif self.chart_type.get() == "line":
                    plt.plot(x, self.df[y_col], label=y_col)
                elif self.chart_type.get() == "scatter":
                    plt.scatter(x, self.df[y_col], label=y_col)
                plt.xlabel(self.selected_x.get())
                plt.ylabel(", ".join(y_cols))
                if self.legend.get():
                    plt.legend()
            filename = os.path.join(self.savepath.get(), "combined_graph.svg")
            plt.savefig(filename, format='svg', bbox_inches='tight')
            plt.close()

        messagebox.showinfo("Generado", "Gráficos guardados exitosamente.")

root = tk.Tk()
app = GraphApp(root)
root.mainloop()
