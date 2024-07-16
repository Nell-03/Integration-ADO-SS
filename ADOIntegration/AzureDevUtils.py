import requests

def get_picklist_field_values(org_url, personal_token, project_name, field_name):
    # Endpoint to get picklist field values
    endpoint = f"{org_url}/{project_name}/_apis/wit/fields/{field_name}/allowedValues?api-version=6.0"

    # Authorization header with personal access token
    headers = {
        "Authorization": f"Basic {personal_token}"
    }

    # Make request to Azure DevOps API
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        data = response.json()
        values = [value['name'] for value in data['items']]
        return values
    else:
        print(f"Failed to retrieve picklist values for field '{field_name}'. Status code: {response.status_code}")
        return []
