# Excel-exporter
Excel-exporter is a boilerplate .NET Framework project designed as a baseline for applications that require Excel file processing capabilities. Developed using ASP.NET, it facilitates the uploading of Excel files, performs predefined logic or calculations, and generates a new Excel file as output.

## Overview
This project was built utilizing the .NET Framework 4.6.1. Although its technology stack may not represent the cutting edge as of now, it serves as a reliable starting point for .NET applications involving Excel operations.

### Key Components
- **Master Page (`site1.master`):** The base page applied to all pages in the project.
- **Default Page (`default.aspx.cs`):** The landing page of the application.
- **Excel.cs:** A class file housing methods to handle basic Excel operations such as creating new Excel files, sheets, reading cells, ranges, etc.

## Features
- Import Excel files.
- Perform specific calculations or logic on the data.
- Export the processed data to a new Excel file.

## Setup & Configuration
This project utilizes Microsoft Office Interop Excel for importing, exporting, and performing cell operations. To enable these functionalities, the `Microsoft.Office.Interop.Excel.dll` needs to be added to the project.

### How to Add Microsoft.Office.Interop.Excel.dll to the Project:
1. Right-click `References` in the project.
2. Under `Assemblies`, search for `Microsoft.Office.Interop.Excel.dll`.
   - If not found, browse to the directory:
     ```plaintext
     C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c
     ```

## Usage
Excel-exporter serves as a foundational project and is intended to be expanded upon to meet specific needs. Users can build upon the existing structure to implement custom logic or calculations and adapt the project to various use cases involving Excel file processing.

## Project Status
This project is currently in a stable state but is not actively maintained. Users are encouraged to update dependencies and refine the codebase to suit contemporary development standards and requirements.

## Disclaimer
Given the timeframe of its development, this project might not adhere to current best practices in software development and may utilize outdated technologies. It is recommended to review and update the codebase as necessary before integrating it into production environments.

## Contact
For further inquiries or clarifications regarding this project, please contact [zell_dev@hotmail.com](mailto:zell_dev@hotmail.com).
