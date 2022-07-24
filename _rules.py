import pandas as pd
from typing import Tuple, Set, List
import warnings


# VARIABLES
DATE = "Assigned Date"
SCREENER = "Screener"

COLOR = "Color"
INFO = "Info"

POSITION_NAME = "Global Position Name"
VOLUNTEER_NAME = "Account Name"
CURRENT_STATUS = "Status Name (Current)"

IS_RECRUITING = "Global Is Recruiting"
INTAKE = "Intake Passed To Region"

BGC = "Last BGC Status"
REGION = "Region Name"

international_prefix = "NHQ:ISD - R&R:"



def closed_and_wrong_region(original_sheet: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, Set[str]]:
    """
    Mark closed positions
    Delete names in wrong region
    Delete names without intake for open positions

    """

    delete_indices = []

    for index, row in original_sheet.iterrows():

        position = row[POSITION_NAME]

        # DELETE
        if not position or pd.isnull(row[VOLUNTEER_NAME]):
            delete_indices.append(index)
            continue

        if row["Opportunity Status"] == "Pending Not In Region":
            delete_indices.append(index)
            continue
        
        # IS OPEN POSITION
        is_closed = row[IS_RECRUITING] == "No"

        # READY TO VOLUNTEER
        new_volunteer = row[CURRENT_STATUS] == "Prospective Volunteer"
        intaked = not pd.isnull(row[INTAKE])
        intake_status = str(row["Intake Progress Status"])
        if intake_status == "Passed to National Headquarters" or intake_status == "Passed to Regional Volunteer Services" or intake_status.startswith("RVS"):
            intaked = True

        not_ready = new_volunteer and not intaked

        if is_closed:

            if not_ready:
                original_sheet.at[index, COLOR] = "ORANGE"
            else:
                original_sheet.at[index, COLOR] = "RED"

            original_sheet.at[index, INFO] = "closed"

        elif not is_closed:

            if not_ready: # applied for open position but not done with intake - ignore for now
                delete_indices.append(index)


    delete_indices = list(set(delete_indices))
    original_sheet = original_sheet.drop(delete_indices)

    return original_sheet


def closed_names(original_sheet: pd.DataFrame) -> pd.DataFrame:
    """
    Rename closed positions

    """
    for index, row in original_sheet.iterrows():
        position = row[POSITION_NAME]

        if row["Color"] == "ORANGE":
            original_sheet.at[index, POSITION_NAME] = position + " (closed for recruitment, not passed to region)"
        if row["Color"] == "RED":
            original_sheet.at[index, POSITION_NAME] = position + " (closed for recruitment)"

    return original_sheet


def auto_assignments(
    original_sheet: pd.DataFrame, rules_sheet) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    """
    Automatically assign volunteers in special conditions to specific screeners
    (ex. weird region, weird status, etc)

    """

    to_kate = "Kate"

    bgc_regions = [
        "Michigan Region", "Greater New York Region", 
        "Eastern New York Region", "Western New York Region", 
        "Kentucky Region"]

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        rules = pd.read_excel(
            rules_sheet, 
            engine="openpyxl")

    for o_index, row in original_sheet.iterrows():
        
        if row[CURRENT_STATUS] == "Biomed Event Based Volunteers":
            original_sheet.at[o_index, INFO] = "BIOMED"
        
        if row[BGC] == "Ready":
            if row["Region Name"] in bgc_regions:
                original_sheet.at[o_index, INFO] = "BGC"
            else:
                original_sheet.at[o_index, SCREENER] = to_kate
                continue
        elif row[BGC] == "Processing":
            original_sheet.at[o_index, INFO] = "BGC"
        
        for _, rule in rules.iterrows():
            if row[rule["Field"]] == rule["Value"]:
                original_sheet.at[o_index, rule["Set"]] = rule["To"]
    
    return original_sheet