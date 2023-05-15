import tkinter as tk
from tkinter import filedialog, messagebox, Menu
import openpyxl
import logging
import requests
import subprocess
import os
from tkinter import ttk
import threading
#developed by Sindre under the MIT license


# set the version and the version URL and the download URL
CURRENT_VERSION = "1.0.1"
VERSION_URL = "https://raw.githubusercontent.com/BeeTwenty/barcode_scanner/master/version.txt"
DOWNLOAD_URL = "https://github.com/BeeTwenty/barcode_scanner/releases/download/Production/BarcodeSetup.exe"


# Add logging to file and console with timestamp and log level and format  (INFO and ERROR)
logging.basicConfig(filename="barcode_log.txt", level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def download_and_install_update():
    # Create a new window for the progress bar
    progress_window = tk.Toplevel()
    progress_window.title("Update Progress")
    progress_window.resizable(False, False)
    progress_window.geometry("400x150")
    progress_window.grab_set()
    progress_window.focus_set()
    
    label = tk.Label(progress_window, text="Downloading update...")
    label.pack(padx=10, pady=10)

    label_file = tk.Label(progress_window, text="Downloading file...")
    label_file.pack(padx=10, pady=5)

    progress_bar = ttk.Progressbar(progress_window, length=300, mode="determinate")
    progress_bar.pack(padx=10, pady=10)


    def download_thread():
        temp_file_path = "BarcodeSetup.exe"
        logging.info("Downloading update...")
        try:
            response = requests.get(DOWNLOAD_URL, stream=True)
            if response.status_code == 200:
                with open(temp_file_path, 'wb') as f:
                    total_size = int(response.headers.get('content-length', 0))
                    block_size = 1024
                    progress = 0


                    for chunk in response.iter_content(chunk_size=1024):
                        f.write(chunk)
                        progress += block_size

                        progress_bar["value"] = (progress / total_size) * 100
                        progress_window.update_idletasks()
                        label_file["text"] = "Downloading file: {:.1f}%".format((progress / total_size) * 100)


                f.close()
                logging.info("Download completed.")
                if os.path.isfile(temp_file_path):
                    progress_window.destroy()
                    subprocess.call([temp_file_path])

                    os.remove(temp_file_path)

                    if messagebox.askyesno("Update", "Update installed successfully. Do you want to exit and restart the application?"):
                        window.quit()
                        logging.info("Update installed successfully. Restarting application.")
                    logging.info("Update installed successfully")
                else:
                    messagebox.showerror("Update Error", "Update installation failed.")
                    logging.error("Update installation failed.")
            else:
                messagebox.showerror("Update Error", "Failed to download update.")
                logging.error("Failed to download update.")

        except requests.exceptions.RequestException as e:
            messagebox.showerror("Update Error", "Error occurred while downloading update: " + str(e))
            logging.error("Error occurred while downloading update: " + str(e))

        except subprocess.CalledProcessError as e:
            messagebox.showerror("Update Error", "Error occurred while installing update: " + str(e))
            logging.error("Error occurred while installing update: " + str(e))

        except Exception as e:
            messagebox.showerror("Update Error", "An error occurred: " + str(e))
            logging.error("An error occurred: " + str(e))
    # Start the download thread
    thread = threading.Thread(target=download_thread)
    thread.start()


def check_updates():
    
    logging.info("Checking for updates...")
    try:
        # Fetch the latest version from the version URL
        response = requests.get(VERSION_URL)
        latest_version = response.text.strip()
        logging.info(f"Latest version: {latest_version}")      
  

        # Compare the current version with the latest version
        if latest_version > CURRENT_VERSION:
            d_response = messagebox.askquestion("Update Available", "A new version ({}) is available. Do you Want to download now?.".format(latest_version))
            if d_response == "yes":
                download_and_install_update() 
            logging.info("Update available. Please update. ( {} )".format(latest_version))
  
        else:
            messagebox.showinfo("Up to Date", "You have the latest version of the program.")
            logging.info("Up to date. ( {} )".format(latest_version))

    except requests.exceptions.RequestException:
        messagebox.showerror("Error", "Failed to check for updates.")
        logging.error("Failed to check for updates.")

def check_updates_at_start():
    logging.info("Barcode Scanner started. Version: {}".format(CURRENT_VERSION))
    logging.info("Checking for updates...")
    try:
        # Fetch the latest version from the version URL
        response = requests.get(VERSION_URL)
        latest_version = response.text.strip()
        logging.info(f"Latest version: {latest_version}")      

        # Compare the current version with the latest version
        if latest_version > CURRENT_VERSION:
            d_response = messagebox.askquestion("Update Available", "A new version ({}) is available. Do you Want to download now?.".format(latest_version))
            if d_response == "yes":
                download_and_install_update()
            logging.info("Update available. Please update. ( {} )".format(latest_version))

        else:
            
            logging.info("Up to date. ( {} )".format(latest_version))

    except requests.exceptions.RequestException:
        messagebox.showerror("Error", "Failed to check for updates.")
        logging.error("Failed to check for updates.")

# Mark barcode in Excel sheet
def mark_barcode_in_excel(barcode, workbook_path, barcode_column):
    try:
        workbook = openpyxl.load_workbook(workbook_path)
        sheet = workbook.active
        logging.info(f"Workbook opened: {workbook_path}")
        barcode_found = False

        # Loop through all cells in the barcode column
        for cell in sheet[barcode_column]:
            if cell.value == barcode:
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", fill_type="solid")  # Mark cell as green
                barcode_found = True
                logging.info(f"Barcode found: {barcode}")

        # Save the workbook
        workbook.save(workbook_path)
        workbook.close()
        logging.info(f"Workbook saved: {workbook_path} and closed.")

        # Show error message if barcode not found or log barcode marked
        if not barcode_found:
            messagebox.showerror("Error", "Barcode not found.")
            logging.error(f"Barcode not found: {barcode}")
        else:
            logging.info(f"Barcode marked: {barcode}")

    # Show error message if workbook not found and log error
    except FileNotFoundError:
        messagebox.showerror("Error", "Workbook not found.")
        logging.error(f"Workbook not found: {workbook_path}")
    
    # Show error message if any other error and log error
    except Exception as e:
        messagebox.showerror("Error", str(e))
        logging.error(f"Error marking barcode: {barcode}. Error: {str(e)}")


# Scan barcode from entry field and mark it in Excel sheet
def scan_barcode(event):
    barcode = barcode_entry.get()
    wb_path = workbook_entry.get()
    bc_column = column_entry.get()
    mark_barcode_in_excel(barcode, wb_path, bc_column)
    barcode_entry.delete(0, tk.END)  # Clear the barcode entry field after scanning
    logging.info(f"Barcode scanned: {barcode}")

# Browse for workbook file
def browse_workbook():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    workbook_entry.delete(0, tk.END)
    workbook_entry.insert(tk.END, file_path)

# Show about window with information about the program
def show_about_window():
    about_text = "Barcode Scanner\n\nVersion: {}\n\nDeveloped by: Sindre\n\nDescription: Enter a barcode to mark it as green in the Excel sheet.\n \n Note: Due to Windows Locking the Excel file when it is open, the program can't run with the file open.".format(CURRENT_VERSION)

    messagebox.showinfo("About", about_text) 
    logging.info("About window opened.")

# Create the main window
window = tk.Tk() # Create the main window
window.title("Barcode Scanner") # Set the window title 
window.geometry("400x300") # Set the window size 
tk.Tk.iconbitmap(window, default="barcode.ico") # Set the window icon
logging.info("Main window created.")
# Create the menu bar
menu = Menu(window) # Create the menu bar
help = Menu(menu, tearoff=0) # Create the Help menu item
help.add_command(label="About", command=show_about_window)
help.add_command(label="Update", command=check_updates) # Add About menu item to Help menu
menu.add_cascade(label="Help", menu=help) # Add Help menu to menu bar
window.config(menu=menu) # Add menu bar to window

# Create the GUI
label_workbook = tk.Label(window, text="Workbook Path:") # Create the workbook path label
label_workbook.pack(pady=10) # Add padding to the label to make it look better

# Create the workbook path entry field
workbook_entry = tk.Entry(window) # Create the workbook path entry field
workbook_entry.pack(padx=5)

# Create the browse button to browse for workbook file
browse_button = tk.Button(window, text="Browse", command=browse_workbook)
browse_button.pack(pady=5)

# Create the barcode column entry field 
label_column = tk.Label(window, text="Barcode Column:")
label_column.pack(pady=10)
column_entry = tk.Entry(window)
column_entry.pack()

# Create the barcode entry field 
label_barcode = tk.Label(window, text="Scan Barcode:")
label_barcode.pack(pady=10)
barcode_entry = tk.Entry(window)
barcode_entry.pack()



# Bind the Return key event to scan_barcode function
barcode_entry.bind("<Return>", scan_barcode)

#check_updates_at_start()
check_updates_at_start()
# Run the main window
window.mainloop()


