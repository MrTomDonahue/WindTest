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
            sheet_names = list(self.data.keys())
            return sheet_names[:3]
        else:
            print("No Excel file loaded.")
            return []

    def get_table_names(self, sheet_name):
        if sheet_name in self.data.keys():
            sheet_data = self.data[sheet_name]
            table1_name = sheet_data.iloc[3, 6]
            table2_name = sheet_data.iloc[4, 0].replace("/", "-")
            return table1_name, table2_name
        else:
            print(f"Sheet '{sheet_name}' does not exist.")

    def get_table(self, sheet_name, table_name):
        if sheet_name in self.data.keys():
            sheet_data = self.data[sheet_name]
            table_names = self.get_table_names(sheet_name)
            if table_name == table_names[0]:
                table = sheet_data.iloc[4:16, 6:13 + 1].reset_index(drop=True)
                table = table.drop(table.columns[[2, 3, 5, 6]], axis=1)  # Delete blank columns
                # table.columns = table.iloc[4, 6:13+1]
                table.columns = table.iloc[0]
                table = table[1:] # remove duplicate column header
                table = table.set_index(table.columns[0])
                return table
            elif table_name == table_names[1]:
                table = sheet_data.iloc[4:16, 0:4 + 1].reset_index(drop=True)
                # main = table.iloc[1, 0:5].tolist()
                main = table.iloc[0].tolist()
                # sub = table.iloc[4, 0:5].tolist()
                sub = table.iloc[1].tolist()
                table.columns = pd.MultiIndex.from_tuples(zip(main, sub))
                return table
            else:
                print("Invalid table name.")
        else:
            print(f"Sheet '{sheet_name}' does not exist.")
