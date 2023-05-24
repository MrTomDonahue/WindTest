import pandas as pd


class ExcelTableReader:
    def __init__(self, url):
        self.url = url
        self.data = None

    def load_excel(self):
        try:
            self.data = pd.read_excel(self.url, sheet_name=None)
            # print("Excel file loaded successfully.")
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")

    def get_sheet_names(self):
        if self.data is not None:
            return list(self.data.keys())
        else:
            print("No Excel file loaded.")
            return []

    def get_table_names(self, sheet_name):
        if sheet_name in self.data.keys():
            sheet_data = self.data[sheet_name]
            table1_name = sheet_data.iloc[3, 6]
            table2_name = sheet_data.iloc[4, 0]
            return table1_name, table2_name
        else:
            print(f"Sheet '{sheet_name}' does not exist.")

    def get_table(self, sheet_name, table_name):
        if sheet_name in self.data.keys():
            sheet_data = self.data[sheet_name]
            table_names = self.get_table_names(sheet_name)
            if table_name == table_names[0]:
                return sheet_data.iloc[5:16, 6:13 + 1]
            elif table_name == table_names[1]:
                return sheet_data.iloc[5:16, 0:4 + 1]
            else:
                print("Invalid table name.")
        else:
            print(f"Sheet '{sheet_name}' does not exist.")
