# PyEmailOutlook
A Python desktop application to send emails through Outlook using Excel-based email details.
This is a simple desktop application built in Python using the `tkinter` library for creating a graphical user interface and the `win32com` library for sending emails through Outlook.

## Features

- Select an Excel file containing email details, including recipients, CC recipients, subject, body, and attachment paths.
- Send emails using the provided email details and Outlook desktop application only.
- Display status messages for successful or failed email sending operations.

## Prerequisites

- Python 3.x installed on your system.
- Required libraries: `tkinter`, `win32com`, and `openpyxl`.

## Setup

1. Ensure that the following directory structure is in one place

2. Open a command prompt or terminal window and navigate to the `PyEmailOutlook` directory:
cd PyEmailOutlook

3. Run the GUI application by executing the following command:
python email_sender_app.py

The application window will open, allowing you to select the Excel file with email details and send the emails using Outlook.
