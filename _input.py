import warnings
import pandas as pd
import os
from datetime import date
from typing import Tuple


# VARIABLES
DATE = "Assigned Date"
SCREENER = "Screener"

COLOR = "Color"
INFO = "Info"

POSITION_NAME = "Global Position Name"
VOLUNTEER_NAME = "Account Name"
CURRENT_STATUS = "Status Name (Current)"
PROGRESS_STATUS = "Progress Status"

REGION = "Region Name"


def open_files(new_data, last_list) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Open files in data folder
    original_sheet = "data*"
    yesterday_sheet = "Screener List*"

    """

    # OPEN FILES
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        original_sheet = pd.read_excel(
            new_data, 
            engine="openpyxl")

    yesterday_sheet = pd.read_excel(last_list)

    return original_sheet, yesterday_sheet


def preliminary_formatting(original_sheet: pd.DataFrame) -> pd.DataFrame:
    """
    Delete unused columns
    Add additional columns

    """

    # FORMAT COLUMNS
    datefields = ["Intake Application Date", "Intake Passed To Region", "Last BGC Created", "Submission Date"]

    for field in datefields:
        try:
            original_sheet[field] = original_sheet[field].dt.strftime('%m/%d/%Y')
        except:
            continue

    original_sheet = original_sheet.drop(["Service Name"], axis=1)

    # ADD COLUMNS
    for index, column in enumerate([DATE, SCREENER, COLOR, INFO]):
        original_sheet.insert(loc=index, column=column, value="")

    return original_sheet


def output_sheets(
    original_sheet: pd.DataFrame, 
    wrong_region: pd.DataFrame) -> None:
    """
    Output sheets

    """

    today = date.today()
    foldername = str(today.strftime("%m-%d-%Y"))

    if not os.path.isdir(foldername):
        os.mkdir(foldername)

    today_sheetname = "LEAD Screener List " + str(today.strftime("%m-%d-%Y")) + ".xlsx"
    wr_sheetname = "WR " + str(today.strftime("%m-%d-%Y")) + ".xlsx"

    # PUT UNASSIGNED AT TOP
    # original_sheet = original_sheet.sort_values(by=[VOLUNTEER_NAME, SCREENER])
    original_sheet.to_excel(os.path.join(foldername, today_sheetname))

    # wrong_region = wrong_region[[REGION, POSITION_NAME, VOLUNTEER_NAME, CURRENT_STATUS, PROGRESS_STATUS]]
    # wrong_region.to_excel(os.path.join(foldername, wr_sheetname))