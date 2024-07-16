import logging
from CreateAzureItems import create_work_item  # Assuming this imports the function correctly
import SS_DataFrame
import SS_Utilities
import Config
import NewProject
# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def mainFunction(projectlist_sheetid):
    try:
        logging.info('Starting row processing...')

        #Get the project Smartsheet Data
        projectlist_df = SS_DataFrame.projectList_dataframe(Config.api_key, projectlist_sheetid) 
        projectNameArray = SS_DataFrame.projectTitle_Array(projectlist_df)  #Project Name Column
        projectDescription_List = SS_DataFrame.projectDescription_Array(projectlist_df) #Project Descr. Column
        projectTemplate_List = SS_DataFrame.ProjectTemplate_Array(projectlist_df) #Project Template Column
        createProject_List = SS_DataFrame.createADO_Array(projectlist_df) #Project Created in ADO Column
        sheetIDArray = SS_DataFrame.sheetID_Array(projectlist_df) #SheetIDs Column
        pushItemsArray = SS_DataFrame.pushItems_Array(projectlist_df) #Push User Stories Column
        testingCompleteList = SS_DataFrame.testingCompleted_Array(projectlist_df) #Testing Completed in ADO column
        
        #Iterate over every row in the Project Smartsheet List
        for index, project in enumerate(projectNameArray):
            
            #Checks if a Project Needs to be created
            if createProject_List[index] == True:
                projectDescription = projectDescription_List[index]
                projectTemplate = projectTemplate_List[index]
                print("Project Name: " + project)
                # Call the function to create a new project
                NewProject.createNew_project(Config.azure_org_url, Config.azure_personal_token, project, projectDescription, projectTemplate)
            else:
                print("No ADO Project created for: " + project)


            #check the status of the push user stories column to push the new user stories to ADO
            if pushItemsArray[index] == "Push User Stories" and createProject_List[index] == True:
                Config.sheet_id = sheetIDArray[index]
                Config.azure_project_name = projectNameArray[index]
                print("Sheet ID: " + str(Config.sheet_id)+ " Project Name: " + Config.azure_project_name)
                workItem_df = SS_DataFrame.smartsheet_to_dataframe(Config.api_key, Config.sheet_id) #Smartsheet Workitems to be pushed to ADO
                adoDataframe = SS_Utilities.get_all_user_stories(Config.azure_org_url, Config.azure_project_name, Config.azure_personal_token)
                alreadyPushed = checkAlreadyPushed(workItem_df, adoDataframe)
                if  alreadyPushed == False:
                    print("Processing Rows...")
                    process_rows(workItem_df, project)
                elif alreadyPushed == True:
                    print("Stories Already Updated")

                SS_Utilities.smartsheet_Update(Config.azure_project_name, Config.sheet_id)

                pushRow_id = SS_Utilities.get_row_id_bycolumn(Config.api_key, projectlist_sheetid, "Project Name", project)
                pushColumn_id = SS_Utilities.get_column_id(Config.api_key, projectlist_sheetid, "Push User Stories")
                SS_Utilities.update_cell(Config.api_key, projectlist_sheetid, pushRow_id, pushColumn_id, "Stories Pushed") #Changes the Value to Stories Pushed in the Project List
            else:
                print("No Stories Pushed for: " + project)

            #check the status of the Testing Completed Column to retrieve the ADO Testing Information
            if createProject_List[index] == True and pushItemsArray[index] == "Stories Pushed" and testingCompleteList[index] == "Get Completed Stories":
                # Get the dataframe from your imported function
                Config.azure_project_name = project
                sheet_id = sheetIDArray[index]
                df = SS_Utilities.get_all_user_stories(Config.azure_org_url, Config.azure_project_name, Config.azure_personal_token)

                if df is not None and not df.empty:
                    SS_Utilities.update_smartsheet_cells(Config.api_key, sheet_id, df)

                pushRow_id = SS_Utilities.get_row_id_bycolumn(Config.api_key, projectlist_sheetid, "Project Name", project)
                pushColumn_id = SS_Utilities.get_column_id(Config.api_key, projectlist_sheetid, "Testing Completed in ADO")
                SS_Utilities.update_cell(Config.api_key, projectlist_sheetid, pushRow_id, pushColumn_id, "Testing Complete") #Changes the Value to Stories Pushed in the Project List
            else:
                print("No Items Imported for: "+ project)

    except Exception as e:
            logging.error(f"An error occurred during processing: {str(e)}")


def process_rows(dataframe, projectName):
    try:
        logging.info('Starting row processing...')

        existing_workitems = SS_Utilities.get_all_work_item_titles(Config.azure_org_url, projectName, Config.azure_personal_token)
        # Iterate over each row in the DataFrame and create work items
        for index, row in dataframe.iterrows():
            try:
                title = row['Title']

                # Skip creating work item if title already exists 
                if title in existing_workitems:
                    logging.info(f"Work item '{title}' already exists. Skipping creation.")
                    continue

                description = row['Description'] if row['Description'] else "No Description"
                area_path = row['AreaPath'] if row['AreaPath'] else projectName
                iteration_path = row['IterationPath'] if row['IterationPath'] else projectName
                work_item_type = row['WorkItemType']
                assigned_to = row['AssignedTo'] if row['AssignedTo'] else " "
                state = row['State']
                acceptance_criteria = row['Acceptance Criteria'] if row['Acceptance Criteria'] else "No Criteria Stated"
                priority = row['Priority'] if row['Priority'] else 1
                due_date = row['DueDate'] if row['DueDate'] else ""
                risk = row['Risk'] if row['Risk'] else " "
                securityGroup = row['Security Group'] if row['Security Group'] else " "
                planType = row['Plan Type'] if row['Plan Type'] else " "

                # Check if any required field is None or empty string after replacement
                if any(value is None or value == '' for value in [title, description, area_path, iteration_path, work_item_type]):
                    raise ValueError("One or more required fields are missing or empty.")

                # Call the function to create the work item
                result = create_work_item(
                    org_url=Config.azure_org_url,
                    project_name=projectName,
                    pat=Config.azure_personal_token,
                    title=title,
                    description=description,
                    area_path=area_path,
                    iteration_path=iteration_path,
                    work_item_type=work_item_type,
                    assigned_to=assigned_to,
                    state=state,
                    acceptance_criteria=acceptance_criteria,
                    priority=priority,
                    due_date=due_date,
                    risk=risk,
                    securityGroup=securityGroup,
                    planType=planType
                )

                if result is not None and 'id' in result:
                    logging.info(f"Successfully created work item '{title}' with ID: {result['id']}")
                else:
                    logging.error(f"Work item creation failed for row {index + 1}: {result}")

            except KeyError as e:
                logging.error(f"Missing field error in row {index + 1}: {str(e)}")
                # Optionally, you can print the row data for further inspection
                logging.error(f"Row data: {row}")

            except ValueError as e:
                logging.error(f"Error creating work item for row {index + 1}: {str(e)}")
                logging.error(f"Row data: {row}")

            except Exception as e:
                logging.error(f"Error creating work item for row {index + 1}: {str(e)}")
                # Optionally, you can print the row data for further inspection
                logging.error(f"Row data: {row}")
    except Exception as e:
        logging.error(f"An error occurred during processing: {str(e)}")


def checkAlreadyPushed(smartsheetDF, adoDataframe):
    smartsheetTitles = set(smartsheetDF['Title'].tolist())
    azureTitles = set(adoDataframe['Title'].tolist())
    
    # Check if all smartsheetTitles are in azureTitles
    for title in smartsheetTitles:
        if title not in azureTitles:
            return False
    
    # If all smartsheetTitles are in azureTitles
    return True




