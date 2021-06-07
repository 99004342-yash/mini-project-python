"""
Program for working with xl sheets.
"""
from ExcelOperation import ExcelOperation

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
