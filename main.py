import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Define the calculation logic
def process_data(inputs):
    i1, i2, i3, i4, i5, i6, i7, i8, i9, i10 = inputs
    
    o1 = i8 - i7
    o2 = i10 - i9
    o3 = i5 - i6
    o4 = o1 / (i3 - i4) if (i3 - i4) != 0 else None
    o5 = o1 / o2 if o2 != 0 else None
    o6 = o4 / i1 if i1 != 0 else None
    
    return [o1, o2, o3, o4, o5, o6]

# Excel read and write functions
def read_from_excel(file_path, sheet_name, start_row, start_col, end_row, end_col):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    data = df.iloc[start_row-1:end_row, start_col-1:end_col].values
    return data

def write_to_excel(file_path, sheet_name, data, start_row, start_col):
    try:
        df = pd.DataFrame(data)
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row-1, startcol=start_col-1, header=False, index=False)
    except PermissionError:
        messagebox.showerror("Error", "Permission denied: Please make sure the file is not open in another application and you have write access.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the UI with Tkinter
class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor")
        self.create_widgets()

    def create_widgets(self):
        # UI Elements
        tk.Label(self.root, text="Input Excel File").grid(row=0, column=0, padx=10, pady=10)
        self.input_file_entry = tk.Entry(self.root, width=40)
        self.input_file_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.load_input_file).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Input Sheet Name").grid(row=1, column=0, padx=10, pady=10)
        self.input_sheet_entry = tk.Entry(self.root, width=20)
        self.input_sheet_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.root, text="Output Excel File").grid(row=2, column=0, padx=10, pady=10)
        self.output_file_entry = tk.Entry(self.root, width=40)
        self.output_file_entry.grid(row=2, column=1, padx=10, pady=10)
        tk.Button(self.root, text="Browse", command=self.load_output_file).grid(row=2, column=2, padx=10, pady=10)

        tk.Label(self.root, text="Output Sheet Name").grid(row=3, column=0, padx=10, pady=10)
        self.output_sheet_entry = tk.Entry(self.root, width=20)
        self.output_sheet_entry.grid(row=3, column=1, padx=10, pady=10)

        tk.Label(self.root, text="Input Range (e.g., A2:J2)").grid(row=4, column=0, padx=10, pady=10)
        self.input_range_entry = tk.Entry(self.root, width=20)
        self.input_range_entry.grid(row=4, column=1, padx=10, pady=10)

        tk.Label(self.root, text="Output Range (e.g., K2)").grid(row=5, column=0, padx=10, pady=10)
        self.output_range_entry = tk.Entry(self.root, width=20)
        self.output_range_entry.grid(row=5, column=1, padx=10, pady=10)

        tk.Button(self.root, text="Process Data", command=self.process_data).grid(row=6, column=1, padx=10, pady=20)

    def load_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.input_file_entry.delete(0, tk.END)
        self.input_file_entry.insert(0, file_path)

    def load_output_file(self):
        file_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx")])
        self.output_file_entry.delete(0, tk.END)
        self.output_file_entry.insert(0, file_path)

    def process_data(self):
        try:
            input_file = self.input_file_entry.get()
            input_sheet = self.input_sheet_entry.get()
            output_file = self.output_file_entry.get()
            output_sheet = self.output_sheet_entry.get()
            input_range = self.input_range_entry.get().split(':')
            output_range = self.output_range_entry.get().split(':')

            # Parse input range
            start_row, start_col = self.parse_range(input_range[0])
            end_row, end_col = self.parse_range(input_range[1])
            # Parse output range
            out_start_row, out_start_col = self.parse_range(output_range[0])

            data = read_from_excel(input_file, sheet_name=input_sheet, start_row=start_row, start_col=start_col, end_row=end_row, end_col=end_col)
            results = [process_data(row) for row in data]
            write_to_excel(output_file, sheet_name=output_sheet, data=results, start_row=out_start_row, start_col=out_start_col)

            messagebox.showinfo("Success", "Data processed and written to the output file successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def parse_range(self, cell_range):
        # Convert cell address like 'A1' to row and column numbers
        col_letter = ''.join(filter(str.isalpha, cell_range))
        row_number = int(''.join(filter(str.isdigit, cell_range)))
        col_number = sum([(ord(letter) - ord('A') + 1) for letter in col_letter])
        return (row_number, col_number)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
