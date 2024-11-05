"""Code using OpenCV"""
import urllib.request
import time
import webbrowser
import threading
import os
from datetime import datetime
from PIL import Image, ImageEnhance
import pandas as pd
import openpyxl
import cv2
import numpy as np
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import xlrd

# PhyPhox server configuration
IPAddress = '172.17.73.61:8080'  # Replace with your PhyPhox IP
num_data = 5  # Number of data pulls in each session
pause_tm = 2  # Time interval between data collections in seconds

# URLs for data management
save_dat = f'http://{IPAddress}/export?format=0'
clear_dat = f'http://{IPAddress}/control?cmd=clear'
start_dat = f'http://{IPAddress}/control?cmd=start'

# Load the image for brightness adjustment
image_path = "image.png"
image = Image.open(image_path)
image_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)  # Convert PIL image to OpenCV format

def lux_to_brightness(lux):
    """Convert Lux value to a brightness factor."""
    return 1 + (lux / 1000)

import urllib.request

def collect_data():
    """Collect data from PhyPhox and save as Excel files with timestamped names."""
    for _ in range(num_data):
        timestamp = datetime.now().strftime("Light %Y-%m-%d_%H-%M-%S")
        file_path = f"./{timestamp}.xlsx"  # Save in the current directory with a timestamped filename

        try:
            # Fetch the data
            response = urllib.request.urlopen(save_dat)
            data = response.read()

            # Debug: Print out the response content
            print("Data content preview:", data[:100])  # Preview first 100 bytes

            # Save the data as an XLSX file
            with open(file_path, 'wb') as f:
                f.write(data)

            print(f"Data collected and saved at {file_path}")
        except Exception as e:
            print(f"Failed to download data: {e}")

        time.sleep(pause_tm)


def clear_and_restart_collection():
    """Clear and restart data collection."""
    try:
        urllib.request.urlopen(clear_dat, timeout=10)
        print("Data cleared.")
        urllib.request.urlopen(start_dat, timeout=10)
        print("Data collection restarted.")
    except urllib.error.URLError as e:
        print("Connection error:", e)

def continuous_data_collection():
    """Thread function for continuous data collection."""
    while True:
        collect_data()
        clear_and_restart_collection()

# def get_lux_values_from_excel(file_path):
#     """Read Lux values from the Excel file."""
#     try:
#         # Check if file exists and can be opened
#         if not os.path.exists(file_path):
#             print(f"File not found: {file_path}")
#             return []
        
#         workbook = openpyxl.load_workbook(file_path)
#         worksheet = workbook.active
        
#         # Print column headers and first few rows for debugging
#         print("Reading Lux data from Excel...")
#         lux_values = []
#         for i, row in enumerate(worksheet.iter_rows(min_row=2), start=2):  # Start from the second row, assuming headers are in the first row
#             row_data = [cell.value for cell in row]
#             print(f"Row {i} data: {row_data}")  # Debug: Print each row's data
            
#             # Assuming Lux is in the second column (adjust index if necessary)
#             lux_value = row[1].value  # Change the index if Lux values are in a different column
#             if lux_value is not None:
#                 lux_values.append(lux_value)
#             else:
#                 print(f"No Lux data in row {i}")
        
#         if not lux_values:
#             print("No Lux values found.")
#         else:
#             print(f"Lux values read successfully: {lux_values}")
#         return lux_values
    
#     except Exception as e:
#         print("Error reading Excel file:", e)
#         return []
def get_lux_values_from_excel(file_path):
    """Read Lux values from an Excel file (either .xls or .xlsx)."""
    try:
        file_extension = os.path.splitext(file_path)[-1].lower()
        if file_extension == '.xls':
            # Use xlrd for .xls files
            df = pd.read_excel(file_path, engine='xlrd')
        elif file_extension == '.xlsx':
            # Use OpenPyXL for .xlsx files
            df = pd.read_excel(file_path)
        else:
            print("Unsupported file format")
            return []
        
        lux_values = df['Illuminance (lx)'].tolist() if 'Illuminance (lx)' in df.columns else []
        return lux_values
    except Exception as e:
        print("Error reading Excel file:", e)
        return []


def adjust_image_brightness_and_display(file_path):
    """Adjust image brightness based on Lux data and display with OpenCV."""
    lux_values = get_lux_values_from_excel(file_path)

    if lux_values:
        print("Lux values read from Excel:", lux_values)  # Print the Lux values to verify correct reading
        for lux in lux_values:
            # Adjust image brightness
            brightness_factor = lux_to_brightness(lux)
            enhancer = ImageEnhance.Brightness(image)
            brightened_image = enhancer.enhance(brightness_factor)
            brightened_image_cv = cv2.cvtColor(np.array(brightened_image), cv2.COLOR_RGB2BGR)

            # Display the adjusted image
            cv2.imshow("Brightness Adjusted by Lux", brightened_image_cv)
            if cv2.waitKey(1000) & 0xFF == ord('q'):  # Display each image for 1 second, press 'q' to exit
                break

class FileEventHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        
        # Print path of newly created file for debugging
        print(f"New file detected: {event.src_path}")

        # Ensure it's an Excel file and try reading data
        if event.src_path.endswith(".xlsx"):
            print(f"Attempting to read data from: {event.src_path}")
            
            # Test read function with debugging information
            lux_values = get_lux_values_from_excel(event.src_path)
            
            if lux_values:
                print(f"Lux values successfully read: {lux_values}")
            else:
                print("No Lux values found or unable to read the file.")
                
            # Now, adjust brightness and update the display
            adjust_image_brightness_and_display(event.src_path)


def main_data_logging():
    urllib.request.urlopen(start_dat)  # Start initial data collection
    data_thread = threading.Thread(target=continuous_data_collection)
    data_thread.daemon = True  # Stops when the main program exits
    data_thread.start()

    # Watch for new Excel files
    event_handler = FileEventHandler()
    observer = Observer()
    observer.schedule(event_handler, ".", recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)  # Periodically check for other tasks
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main_data_logging()
