import tkinter as tk
from tkinter import *
from tkinter import filedialog
import openpyxl
from tkinter import messagebox
from tkinter import Menu
import logging

logging.basicConfig(filename="barcode_log.txt", level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')



def mark_barcode_in_excel(barcode, workbook_path, barcode_column):
    try:
        workbook = openpyxl.load_workbook(workbook_path)
        sheet = workbook.active

        barcode_found = False

        for cell in sheet[barcode_column]:
            if cell.value == barcode:
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", fill_type="solid")  # Mark cell as green
                barcode_found = True

        workbook.save(workbook_path)
        workbook.close()

        if not barcode_found:
            messagebox.showerror("Error", "Barcode not found.")
            logging.error(f"Barcode not found: {barcode}")
        else:
            logging.info(f"Barcode marked: {barcode}")
    except FileNotFoundError:
        messagebox.showerror("Error", "Workbook not found.")
        logging.error(f"Workbook not found: {workbook_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))
        logging.error(f"Error marking barcode: {barcode}. Error: {str(e)}")

def scan_barcode(event):
    barcode = barcode_entry.get()
    wb_path = workbook_entry.get()
    bc_column = column_entry.get()
    mark_barcode_in_excel(barcode, wb_path, bc_column)
    barcode_entry.delete(0, tk.END)  # Clear the barcode entry field after scanning

def browse_workbook():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    workbook_entry.delete(0, tk.END)
    workbook_entry.insert(tk.END, file_path)

def show_about_window():
    about_text = "Barcode Scanner\n\nVersion: 1.0.3\n\nDeveloped by: Sindre\n\nDescription: Enter a barcode to mark it as green in the Excel sheet.\n \n Note: Due to Windows Locking the Excel file when it is open, the program can't run with the file open."

    messagebox.showinfo("About", about_text)

window = tk.Tk()
window.title("Barcode Scanner")
window.geometry("400x300")

menu = Menu(window)
help = Menu(menu, tearoff=0)
help.add_command(label="About", command=show_about_window)
menu.add_cascade(label="Help", menu=help)
window.config(menu=menu)

label_workbook = tk.Label(window, text="Workbook Path:")
label_workbook.pack(pady=10)

workbook_entry = tk.Entry(window)
workbook_entry.pack(padx=5)

browse_button = tk.Button(window, text="Browse", command=browse_workbook)
browse_button.pack(pady=5)

label_column = tk.Label(window, text="Barcode Column:")
label_column.pack(pady=10)
column_entry = tk.Entry(window)
column_entry.pack()

label_barcode = tk.Label(window, text="Scan Barcode:")
label_barcode.pack(pady=10)
barcode_entry = tk.Entry(window)
barcode_entry.pack()

barcode_entry.bind("<Return>", scan_barcode)  # Bind the Return key event to scan_barcode function

window.mainloop()
