import tkinter as tk
from tkinter import filedialog, messagebox
from excel_utils import read_from_excel, write_to_excel
from processor import process_data, calculate_additional_outputs

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

        tk.Label(self.root, text="Hệ số K/Stdev (default 0.03084214)").grid(row=6, column=0, padx=10, pady=10)
        self.base_o9_entry = tk.Entry(self.root, width=20)
        self.base_o9_entry.grid(row=6, column=1, padx=10, pady=10)
        self.base_o9_entry.insert(0, "0.03084214")

        tk.Label(self.root, text="Hệ số kiểm tra/Stdev (default 0.035535571)").grid(row=7, column=0, padx=10, pady=10)
        self.base_o12_entry = tk.Entry(self.root, width=20)
        self.base_o12_entry.grid(row=7, column=1, padx=10, pady=10)
        self.base_o12_entry.insert(0, "0.035535571")

        tk.Label(self.root, text="Sai số (default 0.01)").grid(row=8, column=0, padx=10, pady=10)
        self.margin_error_entry = tk.Entry(self.root, width=20)
        self.margin_error_entry.grid(row=8, column=1, padx=10, pady=10)
        self.margin_error_entry.insert(0, "0.01")

        tk.Button(self.root, text="Bắt đầu", command=self.process_data).grid(row=9, column=1, padx=10, pady=20)

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

            # Get user-defined base values and margin of error
            base_o9 = float(self.base_o9_entry.get())
            base_o12 = float(self.base_o12_entry.get())
            margin_of_error = float(self.margin_error_entry.get())

            data = read_from_excel(input_file, sheet_name=input_sheet, start_row=start_row, start_col=start_col, end_row=end_row, end_col=end_col)
            results = [process_data(row) for row in data]

            additional_outputs, warning_o9, warning_o12 = calculate_additional_outputs(results, base_o9, base_o12, margin_of_error)

            warning_message = []
            if warning_o9:
                warning_message.append("Hệ số K sai số lớn!.")
            if warning_o12:
                warning_message.append("Hệ số kiểm tra sai số lớn!")
            
            if warning_message:
                warning_message.append("Bạn có muốn tiếp tục xuất ra Excel?")
                if not messagebox.askyesno("Warning", " ".join(warning_message)):
                    return

            for result in results:
                result.extend(additional_outputs)
                if warning_o9:
                    result.append("WARNING")
                if warning_o12:
                    result.append("WARNING")

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
