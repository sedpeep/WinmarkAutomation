# Automation Script for Product Management

## Overview
This script automates the process of updating product details on the [Winmark Seller platform](https://winmarkseller.com/products) using data from an Excel sheet. It performs login, navigates through product listings, updates product information, and highlights successfully processed entries in the Excel file.

---

## Features
1. **Automated Login**: Logs into the platform using provided credentials.
2. **Excel Integration**: Reads product IDs from an Excel file (`skus to fix ecom.xlsx`) and highlights processed entries.
3. **Product Updates**:
   - Navigates to the product editing page.
   - Updates product attributes such as weight and checkboxes.
   - Saves changes and verifies completion.
4. **Error Handling**: Skips problematic product IDs and logs errors.
5. **Visual Feedback**: Highlights successfully processed product IDs in yellow.

---

## Prerequisites
1. **Python Libraries**:
   - `selenium`: For browser automation.
   - `openpyxl`: For Excel file manipulation.
2. **Browser Driver**:
   - Google Chrome with the corresponding ChromeDriver installed.
3. **Excel File**:
   - An Excel file named `skus to fix ecom.xlsx` containing product IDs in the first column.

---

## Setup and Usage
1. **Install Dependencies**:
   ```bash
   pip install selenium openpyxl
   ```
2. **Download ChromeDriver**:
   - Ensure it matches your Chrome version.
   - Add it to your system's PATH.

3. **Prepare the Excel File**:
   - Add product IDs in the first column of `skus to fix ecom.xlsx`.

4. **Run the Script**:
   ```bash
   python automation_script.py
   ```

5. **Output**:
   - Successfully processed product IDs will be highlighted in yellow in the Excel file.
   - Errors will be logged in the console.

---

## Key Functionalities

### Chrome Options
- `--disable-blink-features=AutomationControlled`: Prevents detection as a bot.
- `--disable-popup-blocking`: Ensures smooth navigation.

### Product Updates
- Navigates dynamically to product pages.
- Updates attributes like checkboxes and weight fields.
- Uses random weight values between 5 and 7 for variation.

### Excel Updates
- Highlights cells of successfully processed product IDs.
- Saves changes dynamically during processing.

---

## Customization
- **Credentials**: Update login credentials in the `email_elem` and `password_elem` fields.
- **Weight Range**: Modify the `random.randint(5, 7)` range for weight values.
- **Error Logging**: Enhance error handling to log issues into a separate file.

---

## Notes
- Ensure stable internet connectivity for smooth operation.
- Always back up your Excel file before running the script.

---

## Disclaimer
This script is intended for educational and authorized use only. Ensure compliance with the terms and conditions of the platform being accessed. Unauthorized use may lead to account suspension or other consequences.
