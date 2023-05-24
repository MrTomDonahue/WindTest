# NREL Metocean

A showcase of technical skills for a position in OPEX

## Introduction

This is a Python script that allows you to read tables from an Excel file, view them, and save them as CSV or PDF. This script utilizes the pandas library to read Excel files and the ipywidgets library to create graphic user interface (GUI) in Jupyter Notebook.

The data are downloaded from the NREL webpage and represent metoceanic characteristics, averaged based on location. Three locations are represented: West Coast, East Coast and the Gulf of Mexico. The data contain valuable meteorological and ocean condition for offshore wind energy research. You can view and save different  variables of interest as tables for each of the three sites.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Features](#features)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## Prerequisites

- Python 3.x
- pandas library
- ipywidgets library
- pdfkit library
- wkhtmltopdf executable

## Installation

1. Clone the repository:
```python
git clone https://github.com/MrTomDonahue/WindTest.git
```

2. Install the required dependencies:
```python
pip install pandas ipywidgets pdfkit
```

3. Install wkhtmltopdf by following the instruction at: 
https://wkhtmltopdf.org/ 

4.  Set the path to the wkhtmltopdf in the script:
```python
config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
```

## Usage

1. Set the url variable in the script to the URL or local path of the Excel file you want to read.
2. Run the script in a Jupyter Notebook or Python environment.
3. The script will load the Excel file and display a dropdown interface.
4. Select the sheet and table you want to explore.
5. Choose an action from the available options: View, Save as CSV, or Save as PDF.
6. Click the "Execute" button to perform the selected action.
7. The output will be displayed or saved based on the selected action.

## Features
- Load an Excel file and explore its sheets and tables.
- View the selected table within the Jupyter Notebook interface.
- Save the selected table as a CSV file or PDF file.

## Contributing 
Contributions are welcome! If you have any suggestions, improvements, or bug fixes, please open an issue or a pull request on the GitHub repository.

## License
This project is licensed under the MIT License.

## Contact
For any questions, concerns, or feedback, please contact me at: 
tom.s.don2@gmail.com


