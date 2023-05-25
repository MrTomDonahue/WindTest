from ipywidgets import widgets
from IPython.display import display
import pdfkit
from ExcelTableReader import ExcelTableReader


class ExcelDropdown:
    def __init__(self, sheet_names, table_names):
        self.sheet_dropdown = widgets.Dropdown(
            options=sheet_names,
            description="Select sheet:",
        )
        self.table_dropdown = widgets.Dropdown(
            options=table_names,
            description="Select table:",
        )
        self.toggle_buttons = widgets.ToggleButtons(
            options=["View", "Save as CSV", "Save as PDF"],
            button_style="primary",
            description="Action:",
        )
        self.button = widgets.Button(button_style="danger", description="Execute")
        self.output = widgets.Output()

    def create_dropdown(self, url):
        reader = ExcelTableReader(url)
        reader.load_excel()

        def execute_action(btn):
            selected_sheet = self.sheet_dropdown.value
            selected_table = self.table_dropdown.value
            action = self.toggle_buttons.value

            if action == "View":
                table = reader.get_table(selected_sheet, selected_table)
                if table is not None:
                    with self.output:
                        self.output.clear_output()
                        display(table)
            elif action == "Save as CSV":
                table = reader.get_table(selected_sheet, selected_table)
                if table is not None:
                    with self.output:
                        self.output.clear_output()
                        table.to_csv(f"{selected_sheet}_{selected_table}.csv")
                        print(f"Table saved as {selected_sheet}_{selected_table}.csv")
            elif action == "Save as PDF":
                table = reader.get_table(selected_sheet, selected_table)
                if table is not None:
                    with self.output:
                        self.output.clear_output()
                        table_html = table.to_html(index=True)
                        styled_html = f"""
                        <style>
                            table {{
                                font-family: Arial, sans-serif;
                                border-collapse: collapse;
                                width: 100%;
                            }}

                            th, td {{
                                border: 1px solid #ddd;
                                padding: 8px;
                                text-align: left;
                            }}

                            th {{
                                background-color: #f2f2f2;
                            }}
                        </style>

                        {table_html}
                        """

                        pdfkit.from_string(styled_html, f"{selected_sheet}_{selected_table}.pdf")
                        print(f"Table saved as {selected_sheet}_{selected_table}.pdf")

        self.button.on_click(execute_action)
        display(widgets.VBox([self.sheet_dropdown, self.table_dropdown, self.toggle_buttons, self.button, self.output]))
