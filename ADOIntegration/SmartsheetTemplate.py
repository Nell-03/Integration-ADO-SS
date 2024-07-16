import requests
import json
import pandas as pd
import base64
from SS_Utilities import get_column_type

# Azure DevOps constants
azure_org_url = "https://dev.azure.com/lwhite0579/"
azure_personal_token = "cirvdnhouqff5kbwfay36stdbnbfn3me5hgo5sokcgtoddds4nja"
azure_project_name = "Demo"
work_area_id = "1d4f49f9-02b9-4e26-b826-2cdb6195f2a9"

# Smartsheet constants
smartsheet_api_key = "c2crtwuzER4YLTxJtsN6JmApadSr49PmhCk1k"
smartsheet_base_url = "https://api.smartsheet.com/2.0"

# Headers for API requests
smartsheet_headers = {
    "Authorization": f"Bearer {smartsheet_api_key}",
    "Content-Type": "application/json"
}

# List of allowed field names
allowed_fields = [
    "ID", "Title", "Description", "Priority", "AreaPath", "Reason", "IterationPath", "WorkItemType", "AssignedTo",
    "State", "Description", "AcceptanceCriteria", "Priority", "Risk", "DueDate"
]

# Desired column order
desired_column_order = [
    "ID", "WorkItemType", "Title", "State", "Description", "AssignedTo", "Priority", "Reason",
    "AcceptanceCriteria", "Risk", "DueDate", "AreaPath", "IterationPath", 
]

def get_azure_work_items():
    def get_url(org_url, headers, area_id):
        org_resource_areas_url = f"{org_url}/_apis/resourceAreas/{area_id}?api-preview=5.0-preview.1"
        response = requests.get(org_resource_areas_url, headers=headers)
        results = response.json()
        return results.get("locationUrl", org_url) if results else org_url

    token = base64.b64encode(f":{azure_personal_token}".encode("ascii")).decode("ascii")
    headers = {"Authorization": f"Basic {token}"}
    tfs_base_url = get_url(azure_org_url, headers, work_area_id)
    wiql = {"query": f"Select [System.Id] From WorkItems Where [System.TeamProject] = '{azure_project_name}' AND [System.WorkItemType] IN ('User Story', 'Bug')"}
    query_url = f"{tfs_base_url}/{azure_project_name}/_apis/wit/wiql?api-version=5.1"
    response = requests.post(query_url, headers=headers, json=wiql)
    query_result = response.json()

    def process_work_item_fields(fields, work_item_id):
        fields.pop("System.StateChangeDate", None)
        fields_data = {
            field.replace("System.", "")
                  .replace("Microsoft.VSTS.Common.", "")
                  .replace("Microsoft.VSTS.Scheduling.", "")
                  .replace("Microsoft.VSTS.TCM.", ""): value
            for field, value in fields.items() 
            if any(allowed_field in field for allowed_field in allowed_fields)
        }
        fields_data['ID'] = work_item_id  # Add the ID field explicitly
        
        # Ensure 'Created By' column name matches the desired column order
        if 'Created By' in fields_data:
            fields_data['Created By'] = fields_data.pop('Created By')
        
        return fields_data

    work_items = []
    columns = set()

    for work_item_ref in query_result["workItems"]:
        work_item_id = work_item_ref['id']
        work_item_url = f"{tfs_base_url}/_apis/wit/workitems/{work_item_id}?api-version=5.1"
        work_item = requests.get(work_item_url, headers=headers).json()
        work_item_data = process_work_item_fields(work_item['fields'], work_item_id)
        columns.update(work_item_data.keys())
        work_items.append(work_item_data)

    df = pd.DataFrame(work_items, columns=list(columns))
    
    # Only include columns that exist in the DataFrame
    existing_columns = [col for col in desired_column_order if col in df.columns]
    df = df[existing_columns]
    
    return df

def make_unique_column_names(columns):
    seen = {}
    result = []
    for col in columns:
        if col[:50] in seen:
            seen[col[:50]] += 1
            new_col = f"{col[:47]}_{seen[col[:50]]}"
            result.append(new_col)
        else:
            seen[col[:50]] = 0
            result.append(col[:50])
    return result

def create_smartsheet_columns(columns):
    # Ensure 'ID' column is the first column
    if 'ID' in columns:
        columns.remove('ID')
        columns = ['ID'] + columns
    unique_columns = make_unique_column_names(columns)
    return [{"title": col, "primary": (i == 0), **get_column_type(col)} for i, col in enumerate(unique_columns)]


def create_smartsheet_sheet(sheet_name, columns):
    sheet_spec = {
        "name": sheet_name,
        "columns": create_smartsheet_columns(columns)
    }
    response = requests.post(f"{smartsheet_base_url}/sheets", headers=smartsheet_headers, json=sheet_spec)
    return response.json()

def fill_smartsheet_with_data(sheet_id, df):
    df = df.fillna("")  # Replace NaN values with an empty string
    rows = df.to_dict(orient='records')
    for row in rows:
        row_spec = [{"columnId": column_id, "value": value} for column_id, value in row.items()]
        requests.post(f"{smartsheet_base_url}/sheets/{sheet_id}/rows", headers=smartsheet_headers, json=row_spec)

# Main execution
if __name__ == "__main__":
    # Get work item data from Azure DevOps
    work_item_df = get_azure_work_items()
    # Create a new Smartsheet sheet with columns based on work item data
    new_sheet_name = "Work Items from Azure"
    columns = work_item_df.columns.tolist()
    new_sheet_response = create_smartsheet_sheet(new_sheet_name, columns)
    
    if 'result' in new_sheet_response and 'id' in new_sheet_response['result']:
        sheet_id = new_sheet_response['result']['id']
        print(f"Successfully created Smartsheet with ID: {sheet_id}")
        
        # Fill the Smartsheet with data
        fill_smartsheet_with_data(sheet_id, work_item_df)
        print("Successfully filled Smartsheet with data.")
    else:
        print("Failed to create Smartsheet.")
        print(new_sheet_response)
