import shutil
import pandas as pd
import os
from Utilities import *
import xlwings as xw
import logging 


# Functions ------------------------------------------------------------------------------------------------------------------------

def xlsm_editor(path, 
                sheet_name, 
                *cell_values):
    '''
     Edits cell values in a given .xlsm file

     Usage: xlsm_editor(path, sheet_name, (cell_reference,value), (cell_reference,value), ...)

     Expected input: 
        File path, sheet name, cell values
     Expected result: 
        Input values assigned to corresponding cells, in input sheet name, that
        is found at the file path given
    '''
    try:
        with xw.App(visible = False) as app:
            source_sheet, source_book = safe_book_opener(app, 
                                                         path, 
                                                         sheet_name)
            for cell, value in cell_values:
                source_sheet[cell].value = value
            source_book.save() 
    except KeyError as e:
        message = f"Sheet '{sheet_name}' not found in '{path}'"
        error_handling("Error",
                       "Sheet not found in workbook",
                       e, 
                       user_message = message)
        raise SheetError(message) from e
    except Exception as e:
        error_handling("Error",
                       "Unexpected error in xlsm_editor()",
                       e,
                       is_critical = True)
def folder_copier(from_folder, 
                  to_folder):
    '''
    Copies folder "from_folder", and its contents, to "to_folder" file location

    Usage: folder_copier(file path, to_folder_path)

    Expected input:
        Initial location and destination folders as file paths
    Expected result:
        Folder from initial location copied to destination
    '''
    if os.path.exists(from_folder):
        try:
            logging.info(f"Copying {from_folder} to {to_folder}")
            shutil.copytree(from_folder, 
                            to_folder) #copies template folder
        except FileExistsError as e:
            message = f"Folder already exists at:\n '{truncater(to_folder, config['root'])}'"
            error_handling("Warning",
                           message, 
                           e,
                           user_message = message, 
                           file = to_folder)
            raise FileExistsError from e
    else:
        message = f"Error: '{from_folder}' does not exist! \n Check 'config.json' to verify the correct path name has been entered"
        error_handling("Error",
                       f"Folder Copier (): {message}", 
                       "File Not Found",
                       user_message = message)
        raise FileNotFoundError       
def xlsm_reader(file): 
    '''
    Reads in a Excel file and copies it to a dataframe

    Usage: xlsm_reader(file_path)

    Returns: Dictionary containing the last row of the Schedule spreadsheet

    Expected input:
        Excel sheet file path
    Expected result:
        To return the last row of Excel sheet as a dictionary
    '''
    try:
        return pd.read_excel(file,
                             sheet_name = 0).iloc[-1].to_dict()
    except FileNotFoundError as e:
        message = f"File not found at '{file}'"
        error_handling("Error", 
                       f"xlsm_reader({file})", 
                       e, 
                       is_critical = True, 
                       user_message = message)
        raise ExcelError(message) from e
    except Exception as e:
        message = f"Unknown error has occurred while reading\n'{file}'"
        error_handling("Error", 
                       f"xlsm_reader({file})", 
                       e, 
                       is_critical = True, 
                       user_message = message)
        raise ExcelError(message) from e   
def file_path_generator(client_list):
    '''
    Creates a file path based on whether the current "Customer" appears in the client_list.

    Usage: file_path_generator(list)

    Returns: file_name

    Expected input:
        A list of clients
    Expected result:
        A file path
    '''
    try:
        logging.info("Generating file path...")
        client_list_length = len(client_list)
        def plain_text(column_header):
            if not pd.isna(data[column_header]): # Handles blank values, stored in dataframe as 'nan'
                return str(data[column_header]).replace('"',"'").replace("\\","_").replace("/","_").replace("+", "and").replace("LandD", "land")
            else:
                message = f"'{column_header}' value not found please enter a value for it in the schedule"
                raise BlankError(message)
        customer = plain_text("Customer")
        project = plain_text("Project")

        forbidden_characters = ':*<>|'
        if not any(char in customer for char in forbidden_characters) and not any(char in project for char in forbidden_characters):
            schedule_customer = customer.replace(" ","").lower()
            Flag = False
            for i in range(client_list_length):
                client_list_customer = client_list[i].replace(" ","").lower()
                if schedule_customer in client_list_customer or client_list_customer in schedule_customer:
                    logging.info("Client in main_clients list.")
                    client = client_list[i]
                    Flag = True # Needed so that else statement isn't repeatedly run if condition not satisfied
                    break
            if Flag:
                file_name = f'{client.strip()}\\{int(data["Job No."])}_{project.strip()}' 
                return file_name
            else:
                logging.info("Client in main_clients list.")
                file_name = f'MISCELLANEOUS\\{int(data["Job No."])}_{customer.strip()}_{project.strip()}'
                return file_name
        else:
            raise CharacterError(f"Contains {forbidden_characters}")
    except KeyError as e:
        message = "Error reading Main Clients list, check 'config.json'"
        error_handling("Error", 
                       "file_path_generator()", 
                       e,
                       is_critical = True, 
                       user_message = message)
        raise FileGenerationError(message) from e
    except ValueError as e:
        if pd.notna(data["Job No."]):
            message = f"Error using Job Number: '{data['Job No.']}', please enusure that the value only contains numeric characters"
            error_handling("Error", 
                        f"file_path_generator(): {message}", 
                        e,
                        is_critical = True,
                        user_message = message)
        else:
            message = "'Job No.' value is blank, please enter a 'Job No.' value to create a folder."
            error_handling("Error", 
                        f"file_path_generator(): {message}", 
                        e,
                        is_critical = True,
                        user_message = message)
    except BlankError as e:
        error_handling("Warning", 
                        e,
                        is_critical = True,
                        user_message = "Key value missing from Schedule, check that 'Customer' and 'Project' values exist.")
        exit()
    except Exception as e:
        message = "Error generating file path"
        error_handling("Error", 
                       "file_path_generator()", 
                       e,
                       is_critical = True,
                       user_message = message)
        raise FileGenerationError(message) from e

# Main Code -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":

    '''
    The main functioning  of the script has been encapsulated into the main() function for more streamlined error handling.
    '''
    

    def main():
    # Logging ---------------------------------------------------------------------------------------------------------------------------
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The folowing function (defined in the utilities.py file) which is used to configure the log file to store errors. 
        It creates a new folder for it within the working directory if one doesn't already exists. RotatingFileHandler() will 
        create a new log file after the first one reaches as size of > Mb, it does this until there are five full files, then it 
        overwrites the first.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        logger(__file__) 
        
    # File Paths ------------------------------------------------------------------------------------------------------------------------
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The folowing code reads the "config.json" file and stores it as a dictionary that contains file paths and a client list. 

        Then runs the xlsm_reader() function, which returns a dataframe containing the last column from the Schedule as the 
        variable data. Again the main json_reader() function definition has been moved the the utilities.py file.

        Both the variables config and data were set to global to ease their usage through all of this scripts functions.

        Also contained is a pop up to let the user know the process has begun, as the process can take a number of seconds
        to complete.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        global config, data

        config_name = 'config.json'
        config, config_path = json_reader(config_name)
        schedule_path = config["schedule_path"]
        data = xlsm_reader(schedule_path)
        pop_up(f" ⏳ Creating Job Folder for Job: {data['Job No.']}", 
               t = 3000)

        '''
        -------------------------------------------------------------------------------------------------------------------------
        The following is the values in the config dictionary (read in from the .json file) being assigned to a variable for later 
        use. A logging.info() message lets the log know this process has been successful
        -------------------------------------------------------------------------------------------------------------------------
        '''
        main_clients = config["main_clients"]
        schedule_path = config["schedule_path"]
        template_file = config["template_file"] # File Path of Template folder, containing template Control Doc
        destination_directory = config["destination_directory"] # File Path of destination directory (i.e. Projects Folder)
        example_control_doc_name = config["example_control_doc_name"]
        control_doc_name = config["control_doc_name"]
        logging.info(f"'{config_name}' read successfully, '{schedule_path}', '{template_file}', '{destination_directory}', '{example_control_doc_name}', '{control_doc_name}' and '{main_clients}' all read in from '{config_path}'.")

        '''
        -------------------------------------------------------------------------------------------------------------------------
        The function below takes the "main_clients" list from config and sets it all to lowercase and removes any spaces to avoid
        any mismatching errors, it then compares the lowercase version of the most recent customer name to see if they are a main 
        client or not. This then generates a file name, either a main client (without customer name, to be stored in their folder) 
        version or a miscellaneous version (with customer name, to be stored in the miscellaneous folder).
        -------------------------------------------------------------------------------------------------------------------------
        '''
        file_name = file_path_generator(main_clients)
    
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The following code utilises the output of file_path_generator() function to create a file path based on whether a customer 
        appears in the main_clients list, then combines it with a wider "destination directory" to create the complete file path 
        to save the new job folder in. 
        -------------------------------------------------------------------------------------------------------------------------
        '''
        destination_file = f"{destination_directory}{file_name}" # Creating the path for the destination file

        '''
        -------------------------------------------------------------------------------------------------------------------------
        The code below generates the old and new file paths for the control document, so that its name can be changed from 
        "Template control document" to "A-Control Document". It then attempts to copy the template file to the destination, using
        folder_copier(), which contains validation to ensure the file path exists, if it does not an error is logged and a pop up
        message to the user is given.
        Beyond this it attempts to rename the copied "Example Control Document" to the standard "A - Control Document", using
        safe_renamer() (defined in utilities) if the file already carries the same name the user is notified, and if not it is 
        renamed.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        old_ctrl_doc_name  = f"{destination_file}\\{example_control_doc_name}" # Giving the current name (path) of the copied control doc to be changed
        new_ctrl_doc_name = f"{destination_file}\\{control_doc_name}" # Creating the new name (path) of the Control Doc.
        logging.info(f"Copying {template_file} to {destination_file}")
        folder_copier(template_file,
                      destination_file)
        safe_renamer(old_ctrl_doc_name,
                     new_ctrl_doc_name)
        
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The below code utilises the xlsm_editor function to populate the sheets in the new Control Document with the relevant data
        taken from the schedule and stored in the data dictionary.

        Then a pop up is given to let the user know that the process has been successful, and gives them an "open file location"
        button to access the newly generated folder.
        -------------------------------------------------------------------------------------------------------------------------
        '''   
        xlsm_editor(new_ctrl_doc_name, 
                    "Control Doc", 
                    ("D1", data["Customer"]),
                    ("J2", int(data["Job No."])),
                    ("D2", data["Project"]),
                    ("J1", data["Order No."]))
        xlsm_editor(new_ctrl_doc_name,
                    "Del Note",
                    ("F13", data["Del date"]))    
        pop_up(f" ✅ Files successfully created at: \n{truncater(destination_file,destination_directory)}", 
               file = destination_file)

    '''
    -------------------------------------------------------------------------------------------------------------------------
    The below code runs the main() and as previously mentioned allows for critical errors to stop the program and notify the
    user.
    -------------------------------------------------------------------------------------------------------------------------
    '''
    try:
        main()
        logging.info("Run end ================================================================================================================================================================================================\n\n\n")
    except FileNotFoundError as e:
        pass
    except ValueError as e:
        pass
    except KeyError as e:
        pass
    except CharacterError as e:
        message = 'File path contains a forbidden character, please make the "Job No.", "Customer", and "Project" do not contain any of the following characters: \/:*"<>| '
        error_handling("Error", 
                       f"file_path_generator(): {message}",
                       is_critical = True, 
                       user_message = message)
    except BlankError as e:
        pass
    except FileExistsError as e:
        pass
    except Exception as e:
        message = f"Critical Error!\nProcess has been cancelled!"
        error_handling("Error",
                       "main()",
                       e,
                       is_critical = True,
                       user_message = message)
        raise Exception(message) from e

