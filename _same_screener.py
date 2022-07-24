import pandas as pd
from datetime import date
from typing import Dict


# VARIABLES
DATE = "Assigned Date"
SCREENER = "Screener"

COLOR = "Color"
INFO = "Info"

POSITION_NAME = "Global Position Name"
VOLUNTEER_NAME = "Account Name"


def build_screener_map(yesterday_sheet: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    Build dictionary from previous screener list
    {Volunteer Name: {"Screener": Screener Name, Position: Date}}
    """

    screener_map = {}
    for _, row in yesterday_sheet.iterrows():

        name = row[VOLUNTEER_NAME]
        position = row[POSITION_NAME]
        date = row[DATE]
        screener = row[SCREENER]

        if not screener:
            print(name)
            continue
        
        closed_suffix = position.find(" (closed for recruitment")
        if closed_suffix > 0:
            position = position[:closed_suffix]
        
        # add to screener map
        try:
            screener_map[name][position] = date
        except KeyError:
            screener_map[name] = {
                "screener": screener,
                position: date
            }

    return screener_map


def apply_same_screener(
    original_sheet: pd.DataFrame, 
    screener_map: Dict[str, Dict[str, str]]):
    """
    Based on screener map, assign volunteers to the same screeners they had

    """

    today = date.today()
    today_string = str(today.strftime("%-m/%d/%y"))

    for index, row in original_sheet.iterrows():
        name = row[VOLUNTEER_NAME]
        position = row[POSITION_NAME]
        date_set = False

        # if new name in yesterday's list, assign same screener
        # assign old date
        if name in screener_map:
            original_sheet.at[index, SCREENER] = screener_map[name]["screener"]
            if position in screener_map[name]:
                original_sheet.at[index, DATE] = screener_map[name][position]
                date_set = True
        
        # if date unassigned, set as today
        if not date_set:
            original_sheet.at[index, DATE] = today_string
    
    return original_sheet