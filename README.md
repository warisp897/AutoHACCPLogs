Here is the README content formatted in Markdown. You can copy the text below and save it as README.md.

Markdown

# HACCP Log Automation & PDF Generator

This Python script automates the daily processing of Hazard Analysis and Critical Control Point (HACCP) logs. It converts completed Excel logs into dated PDFs for compliance storage and resets the active logs for the next day's use.

## Overview

In a dining unit environment, manual filing and resetting of logs are prone to error. This script performs three critical functions:
1.  **PDF Conversion:** Converts the first sheet of every `.xlsx` file in the "Filled" directory into a PDF.
2.  **Date Stamping:** Appends the current date (`MM-DD-YYYY`) to the filename for archival purposes.
3.  **Log Reset:** Overwrites the completed logs with "Clear" templates, ensuring the unit is ready for the next shift.

## Requirements

* **Operating System:** Windows (Required for `pywin32`).
* **Software:** Microsoft Excel (must be installed on the local machine).
* **Environment:** Python 3.x.
* **Dependencies:**
    ```bash
    pip install pywin32
    ```

## Directory Structure

The script targets specific OneDrive paths synced to the local machine. Ensure the following folders exist within the user's **Virginia Tech OneDrive**:

* `Filled HACCP Logs (FILL IN HERE)`: Source of completed Excel files.
* `Clear HACCP Logs (DO NOT EDIT)`: Source of blank templates.
* `HACCP PDFs`: Destination for generated PDF reports.

## Usage

1.  Complete the Excel HACCP logs throughout the day in the "Filled" folder.
2.  Run the script:
    ```bash
    python "HACCP Gen.py"
    ```
3.  The script will print the name of each file as it processes.
4.  Once finished, the "Filled" folder will be refreshed with blank templates, and the PDFs will be available in the "HACCP PDFs" folder (linked to SharePoint).

## Troubleshooting

* **Excel Hanging:** If the script fails, an invisible instance of Excel may remain open in the background. Open **Task Manager**, find "Microsoft Excel" under **Background Processes**, and select **End Task**.
* **Path Errors:** Ensure the OneDrive folder names exactly match those in the script. The script uses `os.getlogin()` to dynamically find the local user's path.
* **File Permissions:** Do not have the Excel logs open while running the script.

## Script Logic Detail

The script utilizes the `win32com` client to interface directly with the Excel application. It
