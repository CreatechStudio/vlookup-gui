import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import pandas as pd

class VlookupApp:
    def __init__(self, root):
        self.root = root
        self.root.title("VLOOKUP GUI")

        # Variables
        self.origin_file_path = tk.StringVar()
        self.target_file_path = tk.StringVar()
        self.compare_columns = []
        self.compare_options = ['==', '!=', '>', '<']
        self.target_column = tk.StringVar()
        self.fill_column = tk.StringVar()
        self.data_source_column = tk.StringVar()
        self.result_method = tk.StringVar()
        self.result_methods = ['Copy Value', 'Sum', 'Count', 'Unique Count', 'Average', 'Max', 'Min']

        # GUI Components
        ttk.Label(root, text="数据表:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(root, textvariable=self.origin_file_path, state='readonly', width=30).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(root, text="选择文件", command=self.browse_origin_file).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(root, text="查询数据表列:").grid(row=1, column=0, padx=5, pady=5)
        self.compare_column_combobox = ttk.Combobox(root, values=self.compare_columns, state='readonly')
        self.compare_column_combobox.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(root, text="数据源列:").grid(row=2, column=0, padx=5, pady=5)
        self.data_source_column_combobox = ttk.Combobox(root, textvariable=self.data_source_column, state='readonly')
        self.data_source_column_combobox.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(root, text="查询条件:").grid(row=3, column=0, padx=5, pady=5)
        self.compare_option_combobox = ttk.Combobox(root, values=self.compare_options, state='readonly')
        self.compare_option_combobox.grid(row=3, column=1, padx=5, pady=5)
        self.compare_option_combobox.set(self.compare_options[0])

        ttk.Label(root, text="待填写的表:").grid(row=4, column=0, padx=5, pady=5)
        ttk.Entry(root, textvariable=self.target_file_path, state='readonly', width=30).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(root, text="选择文件", command=self.browse_target_file).grid(row=4, column=2, padx=5, pady=5)

        ttk.Label(root, text="待填数据表查询列:").grid(row=5, column=0, padx=5, pady=5)
        self.target_column_combobox = ttk.Combobox(root, textvariable=self.target_column, state='readonly')
        self.target_column_combobox.grid(row=5, column=1, padx=5, pady=5)

        ttk.Label(root, text="待填写列:").grid(row=6, column=0, padx=5, pady=5)
        self.fill_column_combobox = ttk.Combobox(root, textvariable=self.fill_column, state='readonly')
        self.fill_column_combobox.grid(row=6, column=1, padx=5, pady=5)

        ttk.Label(root, text="填写方式:").grid(row=7, column=0, padx=5, pady=5)
        self.result_method_combobox = ttk.Combobox(root, values=self.result_methods, textvariable=self.result_method, state='readonly')
        self.result_method_combobox.grid(row=7, column=1, padx=5, pady=5)
        self.result_method_combobox.set(self.result_methods[0])

        ttk.Button(root, text="运行VLOOK UP", command=self.run_vlookup).grid(row=8, column=1, pady=10)

    def browse_origin_file(self):
        file_path = askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if file_path:
            self.origin_file_path.set(file_path)
            self.load_column_names(file_path, self.compare_column_combobox)
            self.load_column_names(file_path, self.data_source_column_combobox)

    def browse_target_file(self):
        file_path = askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if file_path:
            self.target_file_path.set(file_path)
            self.load_column_names(file_path, self.target_column_combobox)
            self.load_column_names(file_path, self.fill_column_combobox)

    def load_column_names(self, file_path, combobox):
        try:
            df = pd.read_excel(file_path)
            columns = df.columns.tolist()
            combobox['values'] = columns
            if columns:
                combobox.set(columns[0])
        except Exception as e:
            messagebox.showerror("Error", f"Unable to load column names from {file_path}: {e}")

    def run_vlookup(self):
        origin_path = self.origin_file_path.get()
        target_path = self.target_file_path.get()

        if not origin_path or not target_path:
            messagebox.showwarning("Warning", "Please select both origin and target files.")
            return

        fill_column = self.fill_column_combobox.get()
        result_method = self.result_method_combobox.get()

        try:
            origin_df = pd.read_excel(origin_path)
            target_df = pd.read_excel(target_path)

            merged_df = pd.merge(target_df, origin_df, how='left', left_on=self.target_column_combobox.get(), right_on=self.compare_column_combobox.get())

            mask = eval(f"merged_df['{self.compare_column_combobox.get()}'] {self.compare_option_combobox.get()} merged_df['{self.target_column_combobox.get()}']")

            if result_method == 'Copy Value':
                merged_df.loc[mask, fill_column] = merged_df[self.data_source_column_combobox.get()]

            elif result_method == 'Sum':
                merged_df.loc[mask, fill_column] = merged_df.groupby(self.compare_column_combobox.get())[self.data_source_column_combobox.get()].transform('sum')

            elif result_method == 'Count':
                merged_df.loc[mask, fill_column] = merged_df.groupby(self.compare_column_combobox.get())[self.data_source_column_combobox.get()].transform('count')

            elif result_method == 'Unique Count':
                merged_df.loc[mask, fill_column] = merged_df.groupby(self.compare_column_combobox.get())[self.data_source_column_combobox.get()].transform('nunique')

            elif result_method == 'Average':
                merged_df.loc[mask, fill_column] = merged_df.groupby(self.compare_column_combobox.get())[self.data_source_column_combobox.get()].transform('mean')

            elif result_method == 'Max':
                merged_df.loc[mask, fill_column] = merged_df.groupby(self.compare_column_combobox.get())[self.data_source_column_combobox.get()].transform('max')

            elif result_method == 'Min':
                merged_df.loc[mask, fill_column] = merged_df.groupby(self.compare_column_combobox.get())[self.data_source_column_combobox.get()].transform('min')

            # Update the specified columns in target_df with values from merged_df
            target_df.update(merged_df[[self.target_column_combobox.get(), fill_column]])
            
            # Save the result to a new file or modify target.xlsx directly
            target_df.to_excel(target_path.replace('.xlsx', '_result.xlsx'), index=False)

            messagebox.showinfo("Success", "VLOOKUP completed successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Error during VLOOKUP: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VlookupApp(root)
    root.mainloop()
