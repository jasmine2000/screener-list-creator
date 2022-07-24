import warnings
import pandas as pd
from datetime import datetime, date
from typing import List, Dict, Union, Tuple

from yaml import load


# VARIABLES
DATE = "Assigned Date"
SCREENER = "Screener"

COLOR = "Color"
INFO = "Info"

VOLUNTEER_NAME = "Account Name"
PROGRESS_STATUS = "Progress Status"


def load_special_positions(special_positions_sheet) -> Dict[str, Dict[str, Union[str, int]]]:
    screener_limits = {}

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        special_positions = pd.read_excel(
            special_positions_sheet, 
            engine="openpyxl")

    for _, row in special_positions.iterrows():
        note = row["Notes"]
        # position = row["Position name (exact)"]
        names = str(row["Screeners (comma separated)"]).split(",")

        for name in names:
            name = name.strip()
            screener_limits[name] = {"Special": note, "Limit": 0}
    return screener_limits

def create_roster_limits(roster_sheet, special_positions_sheet) -> pd.DataFrame:
    """
    Read Roster sheet and create dictionary in format:
    {Screener Name: 
        {"Limit": limit, 
        "BGC": qualified to run background checks?
        }, 
        ...
    }

    """
    # local constants
    SCREENER_AVAILABLE = "Available to screen"
    LIMIT = "Limit"
    NO_LIST_DAYS = "No List Days"
    BGC_QUALIFIED = "Background Checks"
    WR_QUALIFIED = "Wrong Region"


    screener_limits = load_special_positions(special_positions_sheet)

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        screener_roster = pd.read_excel(
            roster_sheet, 
            engine="openpyxl")
    
    limit_on_total = []

    day = datetime.today().strftime('%A')[0].upper() # first letter of day

    for _, row in screener_roster.iterrows():

        screener_name = row[SCREENER]
        if pd.isnull(screener_name):
            continue

        screener = screener_name.split()[0] # first name

        # unavailable or doesn't take names that day
        is_available = row[SCREENER_AVAILABLE].strip()
        if is_available != "Yes":
            continue
        no_list_days = str(row[NO_LIST_DAYS])
        if day in no_list_days.split('/'):
            continue
        
        # some 'special' 
        if screener not in screener_limits:
            special = None
            if row["Training"] == "Y":
                special = "closed"
            elif str(row[WR_QUALIFIED]).strip() == "Yes":
                special = "WR"
            elif str(row[BGC_QUALIFIED]).strip() == "Yes":
                special = "BGC"
            screener_limits[screener] = {"Special": special}

        if row["Special Only"] == "Y":
            screener_limits[screener]["special_only"] = True

        # limit?
        if not pd.isnull(row[LIMIT]):
            screener_limits[screener]["Limit"] = int(row[LIMIT])
        else:
            screener_limits[screener]["Limit"] = 10
            
        if row["Limit On"] == "week":
            limit_on_total.append(screener)

    return screener_limits, limit_on_total


def get_screener_workload(original_sheet):
    """
    Go through sheet and count total and new names / screener
    Used after "same screener" assignment and to output summary workload
    """
    today = date.today()
    today_string = str(today.strftime("%-m/%d/%y"))
    
    names = {}
    new_names = {}

    for _, row in original_sheet.iterrows():

        screener = row[SCREENER]
        if pd.isnull(screener) or not screener:
            continue

        screener = screener.strip()

        if screener not in names:
            names[screener] = 0
            new_names[screener] = 0
        
        if row[COLOR].lower() == "red" or row[COLOR].lower() == "orange":     # dont count closed positions
            continue
        
        names[screener] += 1

        date_str = row[DATE]

        # date_str = row[DATE].strftime('%-m/%d/%y')

        if (date_str == today_string and pd.isnull(row["Progress Status"])):
            new_names[screener] += 1

    return names, new_names


def create_workload_df(names, new_names):
    """
    Put calculated workload into df so 
    """
    new_dict = {}
    for name, load in names.items():
        new_dict[name] = {"Total": load}
    
    for name, load in new_names.items():
        new_dict[name]["New"] = load

    load_df = pd.DataFrame.from_dict(new_dict, orient='index')
    load_df = load_df.sort_index()
    return load_df


def screener_limit_check(screener_limits):
    # remove screeners from list that have limit <= 0
    to_delete = []
    for name in screener_limits:

        if screener_limits[name]["Limit"] <= 0:
            to_delete.append(name)
    
    for name in to_delete:
        del screener_limits[name]


def adjust_screener_limits(original_sheet: pd.DataFrame, 
    screener_limits: Dict[str, Dict[str, Union[str, int]]],
    weekly_limits: List[str]) -> None:
    """
    Account for screeners that have total limits
    """
    current_workload, _ = get_screener_workload(original_sheet)

    for name in weekly_limits:
        total_limit = screener_limits[name]["Limit"]
        current_names = current_workload[name]

        if total_limit > current_names:
            screener_limits[name]["Limit"] = total_limit - current_names
        else:
            del screener_limits[name]

    screener_limit_check(screener_limits)


def create_special_lists(original_sheet: pd.DataFrame) -> Tuple[Dict[str, str], Dict[str, str], List[int]]:
    assignments = {}
    special_positions = {}
    main = []

    for index, row in original_sheet.iterrows():
        name = row[VOLUNTEER_NAME]
        existing_screener = row[SCREENER]
        special = row[INFO]

        if existing_screener:   # already assigned
            assignments[name] = existing_screener
            continue

        if special:
            if special in special_positions:
                special_positions[special].append(index)
            else:
                special_positions[special] = [index]
        else:
            main.append(index)

    return assignments, special_positions, main


def get_screeners_round(screener_limits, new_names, special = None) -> Tuple[List[str], int]:
    next_round = []
    min_names = float('inf')

    for screener in screener_limits:

        if special:
            for s in special:
                if screener_limits[screener]["Special"] == s:
                    break
            else:
                continue
        else:
            if screener_limits[screener]["Special"] == "closed":
                continue
            if "special_only" in screener_limits[screener]:
                continue

        if screener_limits[screener]["Limit"] > new_names[screener]:
            if new_names[screener] < min_names:
                next_round = [screener]
                min_names = new_names[screener]
            elif new_names[screener] == min_names:
                next_round.append(screener)

    return next_round


def assign_remaining(
    original_sheet: pd.DataFrame, 
    screener_limits: Dict[str, Dict[str, Union[str, int]]],
    weekly_limits: List[str]) -> pd.DataFrame:
    """
    Match screeners and unassigned volunteers

    Algorithm:
    Iterate over sheet
    assign a screener to an unassigned volunteer
    if the screener hasn't hit their limit, add them back into the rotation (at the end)
    """

    adjust_screener_limits(original_sheet, screener_limits, weekly_limits)
    _, new_names = get_screener_workload(original_sheet)
    for screener in screener_limits:
        if screener not in new_names:
            new_names[screener] = 0

    assignments, special_positions, main = create_special_lists(original_sheet)

    for i in range(2):

        if i == 0:
            while_exists = special_positions
        elif i == 1:
            while_exists = main
            current_list = main

        special_position_list = None
        while while_exists:

            if i == 0:
                special_position_list = list(special_positions.keys())

            # get screeners who can be assigned to the special positions
            screener_round = get_screeners_round(
                screener_limits, new_names, special_position_list)

            if not screener_round:
            
                if i == 0:

                    # if closed, just add to main
                    if "closed" in special_positions:
                        main += special_positions["closed"]
                        del special_positions["closed"]

                    # if no screeners for special positions, put in filler and continue
                    for category, index_list in special_positions.items():
                        while index_list:
                            index = index_list.pop(0)
                            original_sheet.at[index, SCREENER] = category.upper()
                        print(while_exists)
                        print(special_positions)
                    special_positions = {}

                # if no screeners for general positions, break
                else:
                    break

            for screener in screener_round:
                
                special_type = None
                if i == 0:
                    special_type = screener_limits[screener]["Special"]
                    try:
                        current_list = special_positions[special_type]
                    except KeyError: # list just got finished
                        continue
                
                assignment = False

                # until this screener gets an assignment or list is empty
                while not assignment and current_list:

                    # get index and name of next name
                    index = current_list.pop(0)
                    next_name = (original_sheet.loc[index])[VOLUNTEER_NAME]

                    # new names might already be assigned to screeners from other position
                    if next_name in assignments:
                        assigned_screener = assignments[next_name]
                    else:
                        assigned_screener = screener
                        assignments[next_name] = assigned_screener
                        assignment = True

                    # update assignment and count
                    original_sheet.at[index, SCREENER] = assigned_screener
                    new_names[assigned_screener] += 1

                # if list empty, remove key
                if i == 0 and not special_positions[special_type]:
                    del special_positions[special_type]

    return original_sheet




