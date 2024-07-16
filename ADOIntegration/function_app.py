import logging
import azure.functions as func
from SmartSheetToAzureMain import mainFunction
import Config

app = func.FunctionApp()

# Import Work items to Azure DevOps
@app.function_name('ADOHTTPUpdateItems')
@app.route(route="integrationroute", methods=['POST'], auth_level=func.AuthLevel.ANONYMOUS)
def handle_request(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request')

    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse(
            "Invalid JSON input",
            status_code=400
        )

    # Extract parameters from the request body
    api_key = req_body.get('smartsheet_api_key')
    projectlist_sheetid = req_body.get('projectlist_sheetid')
    azure_org_url = req_body.get('azure_org_url')
    azure_personal_token = req_body.get('azure_personal_token')

    if not all([api_key, projectlist_sheetid, azure_org_url, azure_personal_token]):
        return func.HttpResponse(
            "Missing required parameters",
            status_code=400
        )

    # Update Config dynamically
    Config.api_key = api_key
    Config.projectlist_sheetid = projectlist_sheetid
    Config.azure_org_url = azure_org_url
    Config.azure_personal_token = azure_personal_token

    try:
        # Call the main function
        mainFunction(projectlist_sheetid)

        return func.HttpResponse(
            body="Processing completed successfully.",
            status_code=200,
            mimetype="text/plain"
        )

    except Exception as e:
        logging.error(f"An error occurred during processing: {str(e)}")
        return func.HttpResponse(
            "Internal Server Error",
            status_code=500
        )
