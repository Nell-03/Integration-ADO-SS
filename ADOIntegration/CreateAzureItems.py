import requests
import base64
from Config import azure_org_url, azure_personal_token, azure_project_name

# Headers for API requests
def get_headers(pat):
    token = base64.b64encode(f":{pat}".encode("ascii")).decode("ascii")
    headers = {"Authorization": f"Basic {token}", "Content-Type": "application/json-patch+json"}
    return headers

def create_work_item(org_url, project_name, pat, title, description, area_path, iteration_path, work_item_type, assigned_to, state, acceptance_criteria, priority, due_date, risk, securityGroup, planType):
    headers = get_headers(pat)
    create_url = f"{org_url}/{project_name}/_apis/wit/workitems/${work_item_type}?api-version=6.0"
    
    payload = [
        {"op": "add", "path": "/fields/System.Title", "value": title},
        {"op": "add", "path": "/fields/System.Description", "value": description},
        {"op": "add", "path": "/fields/System.AreaPath", "value": area_path},
        {"op": "add", "path": "/fields/System.IterationPath", "value": iteration_path},
        {"op": "add", "path": "/fields/System.AssignedTo", "value": assigned_to},
        {"op": "add", "path": "/fields/System.State", "value": state},
        {"op": "add", "path": "/fields/Microsoft.VSTS.Common.AcceptanceCriteria", "value": acceptance_criteria},
        {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority},
        {"op": "add", "path": "/fields/Microsoft.VSTS.Scheduling.DueDate", "value": due_date},
        {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Risk", "value": risk},
        {"op": "add", "path": "/fields/Custom.SecurityGroup", "value": securityGroup}, 
        {"op": "add", "path": "/fields/Custom.PlanType", "value": planType} 
    ]
    
    response = requests.post(create_url, headers=headers, json=payload)
    
    if response.status_code == 200 or response.status_code == 201:
        print("Successfully created work item:")
        print(response.json())
    else:
        print(f"Error creating work item: {response.status_code}")
        print(response.text)

