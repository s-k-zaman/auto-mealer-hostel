import sys

import openpyxl
from pandas.compat import os

from config import (
    failure_string,
    info_string,
    meal_count_sheet_name,
    success_string,
    workbook_name,
)
from excel_utils import get_cell_for_boarder_from_df
from hostel import Hostel
from meal_count_template import init_meal_sheet
from useful_utils import final_cofirm

column_names = [
    "name",
    "total meal",
    "total eggs",
    "guest meals",
    "grand guest meals",
    "guest meal charge",
    "accessories charge",
    "total charge",
    "deposit",
    "due",
    "refund",
]

## if called with init then do this
if len(sys.argv) == 2:
    if sys.argv[1] == "init":
        init_meal_sheet()
    else:
        sys.exit("Usage: python main.py init --> to make a template.")

# else work normally
else:
    workfile_exists = False
    # check if file exist
    if not os.path.exists(workbook_name):
        print(f"{failure_string}file {workbook_name} does not exist!!")
        confirm = final_cofirm("Do you want to create it?")
        if confirm:
            print(f"{info_string}creating {workbook_name}")
            workbook = openpyxl.Workbook()
            init_meal_sheet(workbook)
            print(f"{success_string}{workbook_name} file created.")
            print()
            workfile_exists = True
        else:
            print(f"{failure_string}Can not proceed. Exiting!!")
    else:
        workfile_exists = True
    if workfile_exists:
        print(f"{success_string} Found {workbook_name} file.")
        my_hostel = Hostel(workbook_name)
        my_hostel.main_menu()
