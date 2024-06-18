# Description

This project is a Workers Manager System (WMS) developed to efficiently manage employee data, track salary payments, and facilitate searching and updating records. It consists of multiple windows, each serving different functionalities to streamline the management process effectively.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [File Structure](#file-structure)
- [Publishing](#publishing)
- [Dependencies](#dependencies)
- [Contact Information](#contact-information)
- [Author](#author)

## Features

- **Employee Management Window**: Allows users to view employee workdays, salary deductions, and calculate salaries. Users can input employee data such as name, ID number, joining date, and daily wage. The window also includes search, update, and export functionalities to Excel and Word.

- **Salary Deduction Window**: Enables users to record salary deductions for employees, including details such as name, ID number, date range, and payment date.

- **Interactive Map Window**: Displays an interactive map showing various business locations. Users can search for specific addresses or select locations on the map. It also offers export options to Excel and Word.

## File Structure

- **main.py**: Main Python file containing the application's logic.
- **database.py**: Handles database operations.
- **gui.py**: Defines the graphical user interface elements.
- **mapScreen.py**: Implements the interactive map functionality.
- **advanced.py**: Manages advanced features such as exporting data to Excel and Word.
- **customtkinter.py**: Customizes the appearance of Tkinter widgets.

## Publishing

The application can be published on various platforms such as Windows, macOS, and Linux by packaging it into executable files using tools like PyInstaller or cx_Freeze.

## Dependencies

- Python 3.x
- Tkinter
- pymysql
- openpyxl
- docx
- pptx
- reportlab

## Contact Information

For inquiries or support, please contact [Anis Mahamid] at [email@example.com].

## Author

Developed by Anis Mahamid.
