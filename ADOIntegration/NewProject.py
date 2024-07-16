import requests
import base64
from Config import azure_org_url, azure_personal_token
import json
import logging
def get_headers(pat):
    token = base64.b64encode(f":{pat}".encode("ascii")).decode("ascii")
    headers = {
        "Authorization": f"Basic {token}",
        "Content-Type": "application/json"
    }
    return headers

#Function to create a project
def create_project(org_url, pat, project_name, project_description, template_id):
    headers = get_headers(pat)
    create_url = f"{org_url}/_apis/projects?api-version=6.0"
    payload = {
        "name": project_name,
        "description": project_description,
        "capabilities": {
            "versioncontrol": {
                "sourceControlType": "Git"
            },
            "processTemplate": {
                "templateTypeId": template_id
            }
        }
    }
    
    response = requests.post(create_url, headers=headers, json=payload)
    
    if response.status_code == 200 or response.status_code == 201:
        print("Successfully created project:")
        print(response.json())
    else:
        print(f"Error creating project: {response.status_code}")
        print(response.text)
#Function to get all the projects in an organization
def list_project_names(org_url, pat):
    headers = get_headers(pat)
    projects_url = f"{org_url}/_apis/projects?api-version=6.0"
    
    response = requests.get(projects_url, headers=headers)
    
    if response.status_code == 200:
        projects = response.json()
        project_names = [project['name'] for project in projects['value']]
        return project_names
    else:
        print(f"Error retrieving projects: {response.status_code}")
        print(response.text)
        return None
#Function to get all of the process templates
def get_processes(organization_url, personal_access_token):
    headers = get_headers(personal_access_token)
    base_uri = f"{organization_url}/_apis/process/processes?api-version=6.0"
    
    response = requests.get(base_uri, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error retrieving processes: {response.status_code}")
        print(response.text)
        return None
#Function to get the ID of a process template
def get_template_id(organization_url, personal_access_token, template_name):
    processes = get_processes(organization_url, personal_access_token)
    
    if processes:
        for process in processes['value']:
            if process['name'] == template_name:
                return process['id']
    
    return None
#Function to create a new project using custom parameters
def createNew_project(azure_org_url, azure_personal_token, project_name, project_description, template_name):
    template_id = get_template_id(azure_org_url, azure_personal_token, template_name)
    project_list = list_project_names(azure_org_url, azure_personal_token)
    if template_id:
        print(f"The ID for '{template_name}' is: {template_id}")
        if project_name in project_list:
            logging.info(f"Work item '{project_name}' already exists. Skipping creation.")
        else:
            create_project(azure_org_url, azure_personal_token, project_name, project_description, template_id)
    else:
        raise Exception(f"Template '{template_name}' not found.")
    
if __name__ == "__main__":
        createNew_project(azure_org_url, azure_personal_token, "Demo 3", "Testing 123", "IT Demo Template")
