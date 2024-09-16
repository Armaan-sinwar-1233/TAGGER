# Tag Generation Project

## Project Overview

This project automates the creation of custom clothing tags using product data from an Excel sheet. Each tag includes the product price, description, size, and a barcode. The generated tags are saved in a well-formatted Excel file, ready for printing. The barcodes are created using the `EAN-13` format, and if a barcode already exists, the system reuses it.

## Features
- **Excel Integration**: Reads product data such as item codes, prices, descriptions, and sizes from an input Excel file.
- **Barcode Generation**: Dynamically generates and saves `EAN-13` barcodes for each product, ensuring every item has a unique barcode.
- **Customizable Tags**: Tags are generated with customizable design, including product price, description, size, and barcode.
- **Batch Processing**: The system processes multiple tags at once, allowing for efficient tag generation for large product datasets.

## Tech Stack
- **Python**: Core language for implementing the logic.
- **Pandas**: Used to read and manipulate Excel data.
- **Openpyxl**: Handles Excel file creation and formatting.
- **Pillow (PIL)**: Manages barcode image manipulation.
- **Python-barcode**: Library for generating barcodes in the `EAN-13` format.

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/Armaan-sinwar-1233/Tagger.git
    ```

2. Navigate to the project directory:
    ```bash
    cd Tag-Generation-Project
    ```

3. Install the required Python packages:
    ```bash
    pip install -r requirements.txt
    ```

4. Ensure you have the correct fonts installed (e.g., Calibri) for barcode text rendering, and update the font path in the script if necessary.

## Usage

1. Prepare your product data in an Excel file following the structure in the provided `source_sheet.xlsx` located in the `/data/` folder.
2. Run the script to generate the clothing tags:
    ```bash
    python tagger.py
    ```
3. The tags will be generated and saved as an Excel file in the `/data/` directory.

## Example Output

Here’s an example of a generated barcode included in the Excel tag output:

![Sample Barcode](./images/barcode_example.png)

## Future Enhancements
- Add support for additional barcode formats.
- Improve the customization options for the tag design (e.g., more fonts, colors).
- Provide export options for generating tags in PDF or image formats.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
