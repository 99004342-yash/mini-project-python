

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter


class ExcelOperation:
    """
    ExcelOperation class is used for different Excel operations.

    Parameters:
    name (string): Name of the existing file.
    """

    def __init__(self, name):
        self.name = name
        try:
            self.work_book = load_workbook(name)
        except:
            print(f'File: "{name}" does not exist.')
            quit()

    def get_headers(self, ws):
        """
        get_headers is used for getting the headers of respective file

        Parameters:
        ws : Workbook() class is passed
        """
        return [x.value for x in ws[1]]

    def get_all_ps(self):
        """
        get_all_ps creates an excel file named "otput.xlsx" to show the output.
        """
        ps_no = []
        for ws in self.work_book:
            if not ws.max_row < 2:
                ps_no += [ws[get_column_letter(
                    1) + str(i)].value for i in range(2, ws.max_row)]

        return set(ps_no)

    def get_all_data(self, ps_no):
        char = get_column_letter(1)

        work_book2 = Workbook()

        for ws in self.work_book:
            if ws.max_row < 2:
                continue

            for i in range(2, ws.max_row):
                if ws[char + str(i)].value == ps_no:

                    ws2 = work_book2.create_sheet(ws.title)
                    ws2.append(self.get_headers(ws))
                    ws2.append([x.value for x in ws[i]])
        work_book2.save("output.xlsx")


xls = ExcelOperation('input.xlsx')

while True:

    choice = int(input("""
    ## MENU ##

    Select option: 
    1. Show all PS Numbers
    2. Create XL of selected PS Number.

    """))

    if choice == 1:
        for ps_no_xl in xls.get_all_ps():
            print(ps_no_xl)

    elif choice == 2:
        ps_no_input = (int(input(
            """
        Please enter a valid PS No
        """
        )))
        xls.get_all_data(ps_no_input)
        print("""
        Please check the output.xlsx!

        """)

    else:
        print("""
        Wrong choice
        
        """)
