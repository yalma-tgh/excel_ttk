# Excel Data Entry GUI with Tkinter

This Python script creates a graphical user interface (GUI) using the Tkinter library to facilitate the insertion of data into an Excel spreadsheet. The GUI includes input fields for the user to enter information, a button to insert the data into both the Tkinter interface and an Excel file, and a toggle button to switch between light and dark themes.

## Features

- **Insert Data:** Enter information such as Name, Age, Subscription status, and Employment status. Click the "Insert" button to add a new row to both the Tkinter TreeView widget and the associated Excel file.

- **Toggle Theme:** Use the "Mode" toggle button to switch between light and dark themes for the GUI. The themes are defined using the "forest-light" and "forest-dark" styles.

## Usage

1. Ensure Python is installed on your system.
2. Install the required libraries using `pip install openpyxl`.
3. Run the script using `python main.py`.

## Dependencies

- **Tkinter:** The GUI components are built using Tkinter, the standard GUI toolkit for Python.
- **openpyxl:** Used for reading and writing Excel files.

## Configuration

- The Excel file path is hardcoded as "E:\projects\excel_ttk\people.xlsx". Update the path accordingly in the script.
- The script assumes the existence of "forest-light.tcl" and "forest-dark.tcl" files for themes.

## Notes

- The initial data in the Excel file is loaded into the Tkinter TreeView upon startup.

Feel free to customize the script based on your needs. If you encounter any issues or have questions, please refer to the documentation of Tkinter and openpyxl for further assistance.
