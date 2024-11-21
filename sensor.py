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
import matplotlib.pyplot as plt  # Import for plotting

# PhyPhox server configuration
IPAddress = '172.17.74.162:8080'  # Replace with your PhyPhox IP
num_data = 5
pause_tm = 2

# URLs for data management
save_dat = f'http://{IPAddress}/export?format=0'
clear_dat = f'http://{IPAddress}/control?cmd=clear'
start_dat = f'http://{IPAddress}/control?cmd=start'

# Load the image for brightness adjustment
image_path = "image.png"
image = Image.open(image_path)
image_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)

data_folder = "./Data"
os.makedirs(data_folder, exist_ok=True)

# Initialize metrics and brightness logs
brightness_logs = []  # To store timestamps and brightness factors
data_collection_times = []
file_processing_times = []
image_adjustment_times = []
success_count = {"data_collection": 0, "file_processing": 0, "image_adjustment": 0}
error_count = {"data_collection": 0, "file_processing": 0, "image_adjustment": 0}

def lux_to_brightness(lux):
    """Convert Lux value to a brightness factor."""
    return 1 + (lux / 1000)

def collect_data():
    """Collect data from PhyPhox and save as Excel files in the Data folder."""
    for _ in range(num_data):
        timestamp = datetime.now().strftime("Light %Y-%m-%d_%H-%M-%S")
        file_path = os.path.join(data_folder, f"{timestamp}.xlsx")

        start_time = time.time()
        try:
            response = urllib.request.urlopen(save_dat)
            data = response.read()
            with open(file_path, 'wb') as f:
                f.write(data)

            end_time = time.time()
            data_collection_times.append(end_time - start_time)
            success_count["data_collection"] += 1
            print(f"Data collected and saved at {file_path}")
        except Exception as e:
            error_count["data_collection"] += 1
            print(f"Failed to download data: {e}")

        time.sleep(pause_tm)

def clear_and_restart_collection():
    """Clear and restart data collection."""
    try:
        urllib.request.urlopen(clear_dat, timeout=10)
        urllib.request.urlopen(start_dat, timeout=10)
        print("Data collection restarted.")
    except urllib.error.URLError as e:
        print("Connection error:", e)

def continuous_data_collection():
    """Thread function for continuous data collection."""
    while True:
        collect_data()
        clear_and_restart_collection()

def get_lux_values_from_excel(file_path):
    """Read Lux values from an Excel file."""
    try:
        df = pd.read_excel(file_path)
        lux_values = df['Illuminance (lx)'].tolist() if 'Illuminance (lx)' in df.columns else []
        return lux_values
    except Exception as e:
        error_count["file_processing"] += 1
        print("Error reading Excel file:", e)
        return []

def adjust_image_brightness_and_display(file_path):
    """Adjust image brightness based on Lux data and log values."""
    lux_values = get_lux_values_from_excel(file_path)
    if lux_values:
        for lux in lux_values:
            try:
                brightness_factor = lux_to_brightness(lux)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                brightness_logs.append((timestamp, brightness_factor))  # Log brightness
                
                enhancer = ImageEnhance.Brightness(image)
                brightened_image = enhancer.enhance(brightness_factor)
                brightened_image_cv = cv2.cvtColor(np.array(brightened_image), cv2.COLOR_RGB2BGR)

                cv2.imshow("Brightness Adjusted by Lux", brightened_image_cv)
                if cv2.waitKey(1000) & 0xFF == ord('q'):
                    break
            except Exception as e:
                error_count["image_adjustment"] += 1
                print("Error adjusting image brightness:", e)

class FileEventHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith(".xlsx"):
            adjust_image_brightness_and_display(event.src_path)

def display_metrics():
    """Display collected performance metrics."""
    print("\nPerformance Metrics:")
    print(f"Average Data Collection Time: {sum(data_collection_times) / len(data_collection_times):.2f} seconds")
    print(f"Average File Processing Time: {sum(file_processing_times) / len(file_processing_times):.2f} seconds")
    print(f"Average Image Adjustment Time: {sum(image_adjustment_times) / len(image_adjustment_times):.2f} seconds")
    print(f"Data Collection Success Rate: {100 * success_count['data_collection'] / (success_count['data_collection'] + error_count['data_collection']):.2f}%")
    print(f"File Processing Success Rate: {100 * success_count['file_processing'] / (success_count['file_processing'] + error_count['file_processing']):.2f}%")
    print(f"Image Adjustment Success Rate: {100 * success_count['image_adjustment'] / (success_count['image_adjustment'] + error_count['image_adjustment']):.2f}%")

def plot_brightness_graph():
    """Plot the brightness factor changes over time."""
    timestamps, brightness_factors = zip(*brightness_logs)
    plt.figure(figsize=(10, 6))
    plt.plot(timestamps, brightness_factors, marker='o')
    plt.xlabel('Timestamp')
    plt.ylabel('Brightness Factor')
    plt.title('Brightness Factor Over Time')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def main_data_logging():
    urllib.request.urlopen(start_dat)
    data_thread = threading.Thread(target=continuous_data_collection)
    data_thread.daemon = True
    data_thread.start()

    event_handler = FileEventHandler()
    observer = Observer()
    observer.schedule(event_handler, data_folder, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    cv2.destroyAllWindows()
    plot_brightness_graph()  # Plot the graph after execution
    display_metrics()

if __name__ == "__main__":
    main_data_logging()
