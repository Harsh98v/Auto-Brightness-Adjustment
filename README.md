# Adjusting Luminosity Levels Using Mobile Application

This project implements a chronic data logger that collects light intensity data (Lux) from a PhyPhox application over Wi-Fi. The collected data is then used to adjust the brightness of an image based on environmental conditions. 

## Features

- Continuous data collection from the PhyPhox application.
- Automatic adjustment of image brightness based on Lux values.
- Data saving and clearing operations to manage data size efficiently.
- Threading for simultaneous data logging and brightness adjustment.

## Requirements

- Python 3.x
- Required libraries:
  - `PIL` (Pillow)
  - `OpenCV`
  - `pandas`
  - `urllib`

You can install the required libraries using pip:

```bash
pip install Pillow pandas
```

## Getting Started

### 1. Setting Up PhyPhox

- Install the [PhyPhox app](https://phyphox.org/) on your mobile device.
- Set up the desired sensors (e.g., Light sensor) and ensure the app is running.
- Note the IP address of the PhyPhox server (e.g., `http://10.0.0.236:8080`).

### 2. Modifying the Script

- Update the `phyphox_ip` variable in the script with the correct IP address from your PhyPhox app.
  
```python
phyphox_ip = "http://10.0.0.236:8080"  # Replace with your PhyPhox IP
```

- Provide the path to the Excel file containing the Lux data in the `adjust_image_brightness` function call.

```python
adjust_image_brightness("lux_data.xlsx")
```

### 3. Running the Script

- Run the script:

```bash
python sensor.py
```

The script will start collecting Lux data and adjusting the brightness of the specified image.

## Usage

The script will perform the following operations:

1. Start data collection from the PhyPhox app.
2. Continuously save data every few seconds.
3. Clear the collected data periodically to avoid large file sizes.
4. Adjust the brightness of the specified image based on the Lux values read from the Excel file.

### Customization

You can adjust the following parameters in the script:

- `num_data`: Number of data chunks to collect at once.
- `pause_time`: Time in seconds to wait between data pulls.

## Troubleshooting

- Ensure your computer and the device running PhyPhox are on the same Wi-Fi network.
- Verify that the PhyPhox server is active and reachable using a web browser.
- Check for firewall settings that might block the connection.

## Acknowledgments

- PhyPhox for providing a platform for data collection.
