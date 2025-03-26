# Automated Price List Conversion

## Overview
This project automates the conversion and processing of supplier price lists using a programmed solution. The approach leverages direct data manipulation on a code level, ensuring efficiency and reducing errors compared to manual or macro-based solutions.
The development of this project is a result of my bachelor thesis with the titel

[Potentials of open source attended Robotic Process Automation for small and medium-sized enterprises in procurement logistics:
Use case for the application of AutoHotkey](https://vn4bit.github.io/portfolio/thesisEN.html).

The implementation was carried out using **AutoHotkey V2.0**.

## Features
- Automated extraction of price list data from a PDF file.
- Processing and formatting of the data.
- Conversion to both `.XLSX` and `.CSV` formats.
- Integration with Excel using the COM interface.
- Handling of data inconsistencies and formatting issues.

## Data Processing Details
- **Formatting:**
  - Adjust price formats (thousands separator removal, decimal point standardization).
  - Convert vendor name to supplier ID.
  - Adjust lead time by adding a 3-day delivery buffer.
- **Excel COM Interface:**
  - Uses Excel's COM API for efficient data handling.
  - Saves formatted data into `.XLSX` and `.CSV`.

## Requirements
- Microsoft Excel 2019
- Adobe Acrobat Reader DC
- AutoHotkey (AHK) V2.0

## Usage
1. The PDF price list needs to be in the dedicated folder.
2. Start the AHK script.
