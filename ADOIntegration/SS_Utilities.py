from Config import api_key, sheet_id, azure_org_url, azure_project_name, azure_personal_token, projectlist_sheetid
import smartsheet
import requests
import base64
import pandas as pd
import logging
from pprint import pprint
###################SmartSheet Template CleanUP###############
def get_column_type(column_title):
    # Define a dictionary mapping column titles to Smartsheet column types
    switcher = {
        "ID": "TEXT_NUMBER",
        "WorkItemType": "PICKLIST",
        "Title": "TEXT_NUMBER",
        "State": "PICKLIST",
        "AssignedTo": "CONTACT_LIST",
        "Description": "TEXT_NUMBER",
        "AreaPath": "PICKLIST", 
        "IterationPath": "PICKLIST",
        "Priority": "TEXT_NUMBER", 
        "ActivatedBy": "CONTACT_LIST",  # Use TEXT_NUMBER if CREATED_BY is not valid
        "AcceptanceCriteria": "TEXT_NUMBER",  # Use TEXT_NUMBER if CONTACT_LIST is not valid
        "StoryPoints": "TEXT_NUMBER",
        "Risk": "PICKLIST",
        "StateChangeDate": "DATE",  # Use DATE for date columns
        "DueDate": "DATE"
    }

    # Additional options for picklist columns
    picklist_options = {
        "WorkItemType": ["User Story", "Task", "Bug", "Change Request", "Epic", "Feature", "Issue", "Test Case"],  # Example options for WorkItemType
        "State": ["New", "Active", "Resolved", "In Planning", "Closed"],  # Example options for State
        "AreaPath": [r"Demo"],  # Example options for AreaPath
        "IterationPath": [r"Demo\ERM - UAT Phase 1"],  # Example options for IterationPath
        "Risk": ["1 - High", "2 - Medium", "3 - Low"]  # Example options for Risk
        # Add more options as needed
    }
    
    column_type = switcher.get(column_title, "TEXT_NUMBER")  # Default to "TEXT_NUMBER" if title not found 
    
    if column_type == "PICKLIST":
        options = picklist_options.get(column_title, [])
        if options:
            return {
                "type": "PICKLIST",
                "options": options
            }
    
    return {"type": column_type}

def clean_fields(fields, exclude_substrings):
    cleaned_fields = {}
    for field, value in fields.items():
        # Replace specific substrings
        cleaned_field = (field
                         .replace("System.", "")
                         .replace("Microsoft.VSTS.Common.", "")
                         .replace("Microsoft.VSTS.Scheduling.", "")
                         .replace("Microsoft.VSTS.TCM.", ""))

        # Check if the cleaned field contains any of the exclude substrings
        if not any(exclude in cleaned_field for exclude in exclude_substrings):
            cleaned_fields[cleaned_field] = value

    return cleaned_fields


##################Azure Utilities####################
class WorkItem:
    def __init__(self, id, title, work_item_type, state, description, acceptance_criteria, assigned_to, priority, reason, risk, due_date, area_path, iteration_path):
        self.id = id
        self.title = title
        self.work_item_type = work_item_type
        self.state = state
        self.description = description
        self.acceptance_criteria = acceptance_criteria
        self.assigned_to = assigned_to
        self.priority = priority
        self.reason = reason
        self.risk = risk
        self.due_date = due_date
        self.area_path = area_path
        self.iteration_path = iteration_path

    def __repr__(self):
        return f"WorkItem(id={self.id}, title={self.title})"
# Headers for API requests
def get_headers(pat):
    token = base64.b64encode(f":{pat}".encode("ascii")).decode("ascii")
    headers = {"Authorization": f"Basic {token}", "Content-Type": "application/json"}
    return headers

# Function to fetch all user stories
def get_all_user_stories(org_url, project_name, pat):
    headers = get_headers(pat)
    
    wiql_query = {
        "query": f"Select [System.Id], [System.WorkItemType], [System.Title], [System.State], [System.AssignedTo], [System.CreatedDate], [System.AreaPath], [System.IterationPath] "
                 f"From WorkItems "
                 f"Where [System.WorkItemType] = 'User Story' "
                 f"And [System.TeamProject] = '{project_name}' "
                 f"Order By [System.Id]"
    }
    
    wiql_url = f"{org_url}/{project_name}/_apis/wit/wiql?api-version=6.0"
    response = requests.post(wiql_url, headers=headers, json=wiql_query)
    
    if response.status_code != 200:
        print(f"Error querying work items: {response.status_code}")
        print(response.text)
        return None
    
    work_items = response.json().get("workItems", [])
    
    if not work_items:
        print("No user stories found.")
        return pd.DataFrame()
    
    work_item_details = []
    for item in work_items:
        work_item_id = item["id"]
        work_item_url = f"{org_url}/_apis/wit/workitems/{work_item_id}?api-version=6.0"
        item_response = requests.get(work_item_url, headers=headers)
        
        if item_response.status_code == 200:
            work_item_details.append(item_response.json())
        else:
            print(f"Error fetching work item {work_item_id}: {item_response.status_code}")
            print(item_response.text)
    
    columns = ["ID", "WorkItemType", "Title", "State", "Description", "Acceptance Criteria", "Priority", "Reason","Risk", "Security Group", "Plan Type", "DueDate", "AssignedTo", "AreaPath", "IterationPath"]
    rows = []
    
    for item in work_item_details:
        fields = item["fields"]
        row = {
            "ID": item["id"],
            "WorkItemType": fields.get("System.WorkItemType", ""),
            "Title": fields.get("System.Title", ""),
            "State": fields.get("System.State", ""),
            "Description": fields.get("System.Description", ""),
            "Acceptance Criteria": fields.get("Microsoft.VSTS.Common.AcceptanceCriteria", ""),
            "Priority": fields.get("Microsoft.VSTS.Common.Priority", ""),
            "Reason": fields.get("System.Reason", ""),
            "Risk": fields.get("Microsoft.VSTS.Common.Risk", ""),
            "Security Group": fields.get("Custom.SecurityGroup", ""),
            "Plan Type": fields.get("Custom.PlanType", ""),
            "DueDate": fields.get("Microsoft.VSTS.Scheduling.DueDate", ""),
            "AssignedTo": fields.get("System.AssignedTo", {}).get("uniqueName", ""),
            "AreaPath": fields.get("System.AreaPath", ""),
            "IterationPath": fields.get("System.IterationPath", "")
        }
        rows.append(row)
    
    df = pd.DataFrame(rows, columns=columns)
    return df
#Gets the IDs of specific work items
def get_work_item_id(azure_org_url, azure_project_name, azure_personal_token, work_item_title):
    headers = get_headers(azure_personal_token)
    search_url = f"{azure_org_url}/{azure_project_name}/_apis/wit/wiql?api-version=6.0"
    
    query = {
    "query": f"SELECT [System.Id] FROM workitems WHERE [System.Title] = '{work_item_title}' AND [System.TeamProject] = '{azure_project_name}'"
}


    response = requests.post(search_url, headers=headers, json=query)
    
    if response.status_code == 200:
        work_items = response.json()
        if 'workItems' in work_items and len(work_items['workItems']) > 0:
            work_item_id = work_items['workItems'][0]['id']
            return work_item_id
        else:
            return None
    else:
        print(f"Error retrieving work item: {response.status_code}")
        print(response.text)
        return None

def get_all_work_item_titles(azure_org_url, azure_project_name, azure_personal_token):
    headers = get_headers(azure_personal_token)
    search_url = f"{azure_org_url}/{azure_project_name}/_apis/wit/wiql?api-version=6.0"
    
    # Query to get all work items and their titles
    query = {
    "query": f"SELECT [System.Id], [System.Title] FROM workitems WHERE [System.TeamProject] = '{azure_project_name}'"
    }


    response = requests.post(search_url, headers=headers, json=query)
    
    if response.status_code == 200:
        work_items = response.json()
        titles = []
        if 'workItems' in work_items and len(work_items['workItems']) > 0:
            for item in work_items['workItems']:
                work_item_id = item['id']
                # Retrieve the details of each work item
                details_url = f"{azure_org_url}/{azure_project_name}/_apis/wit/workitems/{work_item_id}?api-version=6.0"
                details_response = requests.get(details_url, headers=headers)
                if details_response.status_code == 200:
                    details = details_response.json()
                    title = details['fields']['System.Title']
                    titles.append(title)
                else:
                    print(f"Error retrieving details for work item {work_item_id}: {details_response.status_code}")
                    print(details_response.text)
        return titles
    else:
        print(f"Error retrieving work items: {response.status_code}")
        print(response.text)
        return []


##########UPDATE SMARTSHEET WITH NEW VALUES########


#Get Smartsheet Columns
def get_all_column_ids(api_key, sheet_id):
    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    
    # Get the sheet details
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    
    # Create a dictionary to map column names to column IDs
    column_ids = {column.title: column.id for column in sheet.columns}
    
    return column_ids

def get_column_id(api_key, sheet_id, column_name):
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    for column in sheet.columns:
        if column.title == column_name:
            return column.id
    return None

#Get Smartsheet Rows
def get_row_ids(api_key, sheet_id, num_rows):
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    
    row_ids = {}
    for index in range(num_rows):
        try:
            row = sheet.rows[index]
            row_ids[index + 1] = row.id
        except IndexError:
            new_row = smartsheet.models.Row()
            new_row.to_top = True  # Add the new row to the top of the sheet
            new_row.cells.append({
                'column_id': sheet.columns[0].id,  # Assuming the first column ID
                'value': f"New Row at Index {index + 1}"
            })
            
            response = smartsheet_client.Sheets.add_rows(sheet_id, [new_row])
            new_row_id = response.data[0].id
            row_ids[index + 1] = new_row_id
    
    return row_ids

def get_row_id_bycolumn(api_key, sheet_id, column_name, target_title):
    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    
    # Get the sheet details
    column_id = get_column_id(api_key,sheet_id, column_name)
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    
    # Search for the row with the specific title
    for row in sheet.rows:
        for cell in row.cells:
            if cell.column_id == column_id and cell.value == target_title:
                return row.id

    raise ValueError(f"No row found with the title '{target_title}'.")



#Update the cells
def update_cell(api_key, sheet_id, row_id, column_id, new_value):
    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    
    # Create the cell object with the new value
    cell = smartsheet.models.Cell()
    cell.column_id = column_id
    cell.value = new_value
    cell.strict = False
    
    # Create the row object with the cell update
    row = smartsheet.models.Row()
    row.id = row_id
    row.cells.append(cell)
    
    # Update the row in the sheet
    updated_row = smartsheet_client.Sheets.update_rows(sheet_id, [row])
    
    return updated_row

def update_specific_row(api_key, sheet_id, row_id, column_id, work_item_title, projectName):
    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    new_value = get_work_item_id(azure_org_url, projectName, azure_personal_token, work_item_title)

    # Build new cell value
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = column_id
    new_cell.value = new_value
    new_cell.strict = False
    
    # Build the row to update
    new_row = smartsheet.models.Row()
    new_row.id = row_id
    new_row.cells.append(new_cell)
    
    # Update the row
    updated_row = smartsheet_client.Sheets.update_rows(
        sheet_id,      # sheet_id
        [new_row]
    )
    
    print(f"Updated row {row_id} with new value: {new_value}")

def update_Smartsheet_IDs(workItem_title, sheet_id, projectName):
    try:

        column_id = get_column_id(api_key, sheet_id, "ID")  # get ID, Column ID
        specific_row_id = get_row_id_bycolumn(api_key, sheet_id, "Title",workItem_title)  # Get the row of an item from its title

        if specific_row_id is not None:
            update_specific_row(api_key, sheet_id, specific_row_id, column_id, workItem_title, projectName)  # Update the Item's ID in Smartsheet
        else:
            logging.warning(f"Row with title '{workItem_title}' not found in Smartsheet.")

    except Exception as e:
        logging.error(f"An error occurred during Smartsheet update: {str(e)}")

def update_pushcolumns(api_key, sheet_id, column_name, row_id):
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    
    # Retrieve the sheet
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    
    # Get column ID based on column name
    column_id = get_column_id(api_key, sheet_id, column_name)
    
    if column_id:
        # Find the specific row by row_id
        update_row = None
        for row in sheet.rows:
            if row.id == row_id:
                update_row = row
                break
        
        if update_row:
            # Update cell in the specific column and row
            for cell in update_row.cells:
                if cell.column_id == column_id:
                    cell.value = "Stories Pushed"  # Replace with appropriate value based on column_name
                    break
            
            # Send update request for the specific row
            response = smartsheet_client.Sheets.update_rows(sheet_id, [update_row])
            print(f"Updated column '{column_name}' for row ID '{row_id}' successfully.")
        else:
            print(f"Row ID '{row_id}' not found in the sheet.")
    else:
        print(f"Column '{column_name}' not found.")





############# Main Methods ##################
#Used to update only the ID Column of a smartsheet after items have been uploaded
def smartsheet_Update(projectName, sheet_id):
    try:
        az_titles = get_all_work_item_titles(azure_org_url, projectName, azure_personal_token)
        print (az_titles)
        # Update Smartsheet IDs for each title
        for title in az_titles:
            update_Smartsheet_IDs(title, sheet_id, projectName)


    except Exception as e:
        logging.error(f"An error occurred during processing: {str(e)}")

#Used to update an entire Smartsheet with Workitems
def update_smartsheet_cells(api_key, sheet_id, df):
    smartsheet_client = smartsheet.Smartsheet(api_key)
    smartsheet_client.errors_as_exceptions(True)
    num_rows = len(df)
    row_ids = get_row_ids(api_key, sheet_id, num_rows)
    
    for index, row in df.iterrows():
        row_id = row_ids.get(index + 1)
        
        if row_id is not None:
            row_to_update = smartsheet.models.Row()
            row_to_update.id = row_id
            cells = []
            processed_column_ids = set()  # Track processed column IDs to prevent duplicates
            for column_name in df.columns:
                column_id = get_column_id(api_key, sheet_id, column_name)
                if column_id is not None and column_id not in processed_column_ids:
                    cell = smartsheet.models.Cell()
                    cell.column_id = column_id
                    cell.value = str(row[column_name])
                    cells.append(cell)
                    processed_column_ids.add(int(column_id))  # Mark column ID as processed
                else:
                    print(f"Column '{column_name}' not found in Smartsheet or already processed.")
            
            row_to_update.cells = cells
            updated_row = smartsheet_client.Sheets.update_rows(sheet_id, [row_to_update])
            print(f"Updated row: {updated_row.data[0].id}")
        else:
            print(f"Row for identifier '{row['ID']}' not found in Smartsheet.")


    
    
    