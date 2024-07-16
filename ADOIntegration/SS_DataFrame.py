import pandas as pd
import smartsheet
import numpy as np
from Config import api_key, sheet_id, projectlist_sheetid
import Config
def smartsheet_to_dataframe(api_key, sheet_id):
    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)

    # Load the sheet
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

    # Extract column names
    column_names = [column.title for column in sheet.columns]

    # Extract rows
    rows = []
    for row in sheet.rows:
        row_data = [cell.value for cell in row.cells]
        
        # Check if the row has either ID or WorkItemType
        if 'ID' in column_names and 'WorkItemType' in column_names:
            if row_data[column_names.index('ID')] or row_data[column_names.index('WorkItemType')]:
                rows.append(row_data)
        elif 'ID' in column_names:
            if row_data[column_names.index('ID')]:
                rows.append(row_data)
        elif 'WorkItemType' in column_names:
            if row_data[column_names.index('WorkItemType')]:
                rows.append(row_data)

    # Convert to Pandas DataFrame
    df = pd.DataFrame(rows, columns=column_names)

    return df

def smartsheet_Titlesdf(df):
    # Extract values of 'Push User Stories' 
    titleArray = df['Title'].tolist()
    
    return titleArray

def projectList_dataframe(api_key, projectlist_sheetid):
    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)

    # Load the sheet
    sheet = smartsheet_client.Sheets.get_sheet(projectlist_sheetid)

    # Extract column names
    column_names = [column.title for column in sheet.columns]

    # Extract rows
    rows = []
    for row in sheet.rows:
        row_data = [cell.value for cell in row.cells]
        
        # Check if the row has either Project Name, SheetID, Push User Stories, or Project Created in ADO
        if 'Project Name' in column_names and 'SheetID' in column_names:
            if row_data[column_names.index('Project Name')] or row_data[column_names.index('SheetID')]:
                rows.append(row_data)
        elif 'Project Name' in column_names:
            if row_data[column_names.index('Project Name')]:
                rows.append(row_data)
        elif 'SheetID' in column_names:
            if row_data[column_names.index('SheetID')]:
                rows.append(row_data)

    # Convert to Pandas DataFrame
    df = pd.DataFrame(rows, columns=column_names)

    return df

def pushItems_Array(df):
    # Extract values of 'Push User Stories' 
    push_user_stories = df['Push User Stories'].tolist()
    
    return push_user_stories 

def projectTitle_Array(df):
    # Extract values of 'Push User Stories' 
    titleArray = df['Project Name'].tolist()
    
    return titleArray

def sheetID_Array(df):
    # Fill NaN values with a placeholder (if applicable) or drop them
    df['SheetID'] = df['SheetID'].fillna(0)  # You can change 0 to any placeholder or use dropna() if appropriate
    
    # Ensure all values are integers
    sheetIDS = df['SheetID'].astype(int).to_numpy()
    
    return sheetIDS

def projectDescription_Array(df):
    # Extract values of 'Push User Stories' 
    projectDescrition_List = df['Project Description'].tolist()
    
    return projectDescrition_List

def ProjectTemplate_Array(df):
    # Extract values of 'Push User Stories' 
    projectTemplate_Array = df['Project Template Name'].tolist()
    
    return projectTemplate_Array

def createADO_Array(df):
    # Extract values of 'Push User Stories' 
    createProjectList = df['Project Created in ADO'].tolist()
    
    return createProjectList

def testingCompleted_Array(df):
    # Extract values of 'Push User Stories' 
    testsCompleteList = df['Testing Completed in ADO'].tolist()
    
    return testsCompleteList


