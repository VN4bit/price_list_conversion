# Automated Price List Conversion

This project automates the conversion and processing of supplier price lists using a programmed solution. The approach leverages direct data manipulation on a code level, ensuring efficiency and reducing errors compared to manual or macro-based solutions.

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

## Challenges and Considerations
- The Excel COM interface documentation is sparse, requiring extensive testing.
- Copying data directly from PDF introduces potential formatting issues.

