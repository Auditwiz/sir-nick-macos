import customtkinter as ctk
import openpyxl
from tkinter import filedialog, messagebox
import os

class ExcelCleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("Sir Nick - Excel Formatter")
        self.geometry("400x240")
        ctk.set_appearance_mode("System")  # Follows macOS dark/light mode
        ctk.set_default_color_theme("blue")

        # UI Elements
        self.label = ctk.CTkLabel(self, text="Excel Number Formatter", font=("Arial", 20))
        self.label.pack(pady=20)

        self.btn_select = ctk.CTkButton(self, text="Select & Process File", command=self.process_logic)
        self.btn_select.pack(pady=20)

        self.status_label = ctk.CTkLabel(self, text="Ready", font=("Arial", 12))
        self.status_label.pack(pady=10)

    def process_logic(self):
        # 1. File Selection
        input_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not input_path:
            return

        try:
            self.status_label.configure(text="Processing...")
            self.update_idletasks()

            # 2. Excel Logic (Optimized to prevent XML corruption)
            wb = openpyxl.load_workbook(input_path)
            ws = wb.worksheets[0]
            clean_number_fmt = '#,##0.00'

            for row in ws.iter_rows():
                for cell in row:
                    if cell.hyperlink or cell.value is None:
                        continue
                    try:
                        raw_val = str(cell.value).replace(',', '').strip()
                        if raw_val.replace('.', '', 1).replace('-', '', 1).isdigit():
                            cell.value = float(raw_val)
                            cell.number_format = clean_number_fmt
                    except:
                        continue

            # 3. Save Selection
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"Cleaned_{os.path.basename(input_path)}"
            )

            if output_path:
                wb.save(output_path)
                messagebox.showinfo("Success", f"File saved to:\n{output_path}")
                self.status_label.configure(text="Success!")
            else:
                self.status_label.configure(text="Save Cancelled")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.configure(text="Error occurred")

if __name__ == "__main__":
    app = ExcelCleanerApp()
    app.mainloop()

# Made with Bob
