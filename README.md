# Comparison of Delivery Dates

## Overview
This project automates the comparison of delivery dates for components in an ERP simulation using Excel. 
The process determines the latest delivery date of two components and enters it into the corresponding column.
The development of this project is a result of my bachelor thesis with the titel

**[Potentials of open source attended Robotic Process Automation for small and
medium-sized enterprises in procurement logistics:<br> Use case for the application of AutoHotkey](https://vn4bit.github.io/portfolio/thesisEN.html)**.

The implementation was carried out using **AutoHotkey V2.0**.

![alt text](https://github.com/VN4bit/2_dates_comparison/blob/main/gifs/excel.gif)

## Functionality
The robot processes the ERP information (simulated with Excel) and compares the delivery dates of two components. The later date is automatically entered into the **Actual Delivery Date** column.

## Implementation
### Custom Programming
**Development time:** 5 hours

**Procedure:**
1. Search for and store the delivery date for Component 1 using **Image Search**.
2. Search for and store the delivery date for Component 2 using **Image Search**.
3. Convert both values into numbers and compare them.
4. Store the later date in a variable.
5. Locate the **Actual Delivery Date** column and insert the date.
6. Check whether additional data exists below:
   - If yes: Continue.
   - If no: End the process.

## Requirements
- Microsoft Excel (for simulating the ERP file)
- AutoHotkey (AHK) V2.0

## Usage
1. The **Excel file** must be open.
2. Start the AHK script.

## License
This project is open-source and available under the [MIT License](LICENSE).
