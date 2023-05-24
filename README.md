# NREL Metocean
A showcase of technical skills for a position in OPEX
This is a Python script that allows you to read tables from an Excel file, view them, and save them as CSV or PDF.

## Prerequisites

- Python 3.x
- pandas
- ipywidgets
- pdfkit
- wkhtmltopdf

## Installation

1. Clone the repository:
git clone https://github.com/MrTomDonahue/WindTest.git

2. Install the required dependencies:
pip install pandas ipywidgets pdfkit


3. Set the path to the wkhtmltopdf executable in the script:

```python
# Set the path to the wkhtmltopdf executable
config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')

## Usage

1. Load the Excel file:
url = "https://www.nrel.gov/wind/nwtc/assets/downloads/MetOcean/DistributionParameters.xlsx"
reader = ExcelTableReader(url)
reader.load_excel()

2. Get the available sheet names:
sheet_names = reader.get_sheet_names()

3. Create an ExcelDropdown widget:
table_names = reader.get_table_names(sheet_names[0])
dropdown = ExcelDropdown(sheet_names, table_names)
dropdown.create_dropdown()

4. Select a sheet, table, and action from the dropdown menus, and click the "Execute" button.

## License
This project is licensed under the MIT License.

