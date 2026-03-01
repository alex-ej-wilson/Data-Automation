import xlwings as xw
import pandas as pd
from Utilities import *
import logging 


# Functions --------------------------------------------------------------------------------------------------------------------------------

def sheet_copier(source_sheet, 
                 destination, 
                 destination_book):
    '''
    Copies a source_sheet to destination_workbook, and sets it to hidden, if the sheet already exists it is replaced with a new copy

    Usage: sheet_copier(source_sheet, destination_path ,destination_workbook)

    Expected input:
        An Excel sheet to be copied, a file path to a destination workbook, the destination workbook
    Expected result:
        The given sheet has been copied into the destination workbook, and is hidden
    '''
    try:
        logging.info(f"Copying '{source_sheet}' to '{destination}'")
        sheet_name = source_sheet.name
        validate_sheet_name(destination_book, 
                            sheet_name)
        sheet_to_delete = destination_book.sheets[sheet_name]
        sheet_to_delete.delete()
        source_sheet.copy(after=destination_book.sheets[0])
        copied_sheet = destination_book.sheets[sheet_name]
        copied_sheet.visible = False
        destination_book.save()
        return True
    except KeyError as e:
        message = f"Sheet '{source_sheet}' not found"
        error_handling("Error", 
                       "Sheet copier()", 
                       e,
                       is_critical = True, 
                       user_message = message)
        raise ExcelError(message) from e
    except Exception as e:
        message = f"An unknown error has occurred while trying to copy the '{source_sheet}' sheet"
        error_handling("Error", 
                       "sheet_copier()", 
                       e,
                       is_critical = "True", 
                       user_message = message)
def drop_down(sheet, 
              formula, 
              given_range):
    '''
    Creates drop downs in the given sheet containing values equal to the given formula, across the given range.

    Usage: drop_down(sheet, formual, range)

    Expected input:
        An Excel sheet, a validation formula, a continuous range of cells to validate
    Expected output:
        Drop down menus in the given range in the given sheet containing the given formula
    '''
    try:
        cell_range = sheet.range(given_range)
        cell_range.api.Validation.Delete()
        cell_range.api.Validation.Add(Type = 3, Formula1 = formula) #IMPORTANT
        cell_range.api.Validation.IgnoreBlank = True #IMPORTANT
        cell_range.api.Validation.ShowError = False #IMPORTANT
    except ValueError as e:
        message = f"error creating drop downs for {cell_range}"
        error_handling("Error", 
                       message, 
                       e,
                       is_critical = True, 
                       user_message = message)
    except Exception as e:
        message = f"An unknown error has occurred while trying to create drop down menus for '{sheet}'"
        error_handling("Error", 
                       message, 
                       "sheet_copier()",
                       is_critical = True, 
                       user_message = message)
        raise ExcelError(f"An unknown error has occurred while trying to create drop down menus for '{sheet}'") from e
def DB_reader(path_, 
              column_name, 
              sheet_num = 0):
    '''
    Obtains the length of the column given in the 0th (by default) sheet of the work book located at the given path.

    Usage: DB_reader(path, column_name, sheet_number(optional))

    Returns: a string that is a cell range in the Excel format that will encompass all data in the chosen column

    Expected input:
        File path to Excel workbook, a column name, (optional) a sheet number
    Expected output:
        Returns an Excel formula corresponding to the number of rows data in 'Item Database'
    '''
    try:
        logging.info("Reading Item Description Column")
        df = pd.read_excel(path_, 
                           sheet_name = sheet_num)
        itdesc = df[column_name]
        id_arr = itdesc.to_numpy()
        length = len(id_arr)
        formula = f"='Item Database'!$B$2:$B{length+1}"#verify sheet name if encountering issues
        return str(formula)
    except FileNotFoundError as e:
        message = f"File not found at '{path_}'"
        error_handling("Error", 
                       f"{message} at DB_reader()", 
                       e,
                       is_critical = True, 
                       user_message = message)
        raise FileNotFoundError(f"File not found at '{path_}'")
    except ValueError as e:
        message = f"An error has occurred while trying to read the Item Database at {path_}"
        error_handling("Error", 
                       f"{message} at DB_reader()",
                       is_critical = True, 
                       user_message = message)
        raise ExcelError(message) from e
    except Exception as e:
        message = f"An unknown error has occurred while trying to read the Item Database at {path_}"
        error_handling("Error", 
                       "DB_reader()", 
                       is_critical = True,
                       user_message = message)
        raise ExcelError(message) from e

# Main Code -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    
    '''
    The main functioning  of the script has been encapsulated into the main() function for more streamlined error handling.
    '''
    

    def main():
        
        # Logging ---------------------------------------------------------------------------------------------------------------------------
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The code below contains a pop up to let the user know the process has begun, as the process can take a number of seconds
        to complete.
        
        Also is the function (defined in the utilities.py file) that is used to configure the log file to store errors. 
        It creates a new folder for it within the working directory if one doesn't already exists. RotatingFileHandler() will 
        create a new log file after the first one reaches as size of > Mb, it does this until there are five full files, then it 
        overwrites the first.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        pop_up(" ⏳ Updating Template Control Document", 
               t = 3000)
        logger(__file__) 
        app_active = xw.apps.active
        for book in app_active.books:
            if book.name == "A-Control Doc.xlsm":
                message = "Example Control Document open, please close before attempting to update drop down menus"
                error_handling("Warning",
                                "Example Control Document open", 
                                is_critical = True, 
                                user_message = message)
                raise WorkBookOpenError
        
        # Paths ----------------------------------------------------------------------------------------------------------------------------------
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The folowing code reads the "config.json" file and stores it as a dictionary that contains file paths. 
        Again the main json_reader() function definition has been moved the the utilities.py file.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        config_name = 'config.json'
        config, config_path = json_reader(config_name)
        path_itdb = config["path_itdb"]
        path_ctrl = config["path_ctrl"]
        logging.info(f"'{config_name}' read successfully.\n '{path_itdb}','{path_ctrl}' read in from '{config_path}'.")
        process_successful = "Excel opening"
        logging.info(f"Process stage completed: {process_successful}")

        '''
        -------------------------------------------------------------------------------------------------------------------------
        The following try and except block tries to open a new instance of Excel, that runs in the background. Then the 
        'Item Database' and a the 'Template Control Document' sheets are accessed using safe_book_opener() (defined in utilities.py) 
        which contains its own validation and error handling. 
        -------------------------------------------------------------------------------------------------------------------------
        '''
        try:
            logging.info("Opening new instance of Excel")
            app = xw.App(visible = False)
        except Exception as e:
            message = f"An unknown error has occurred while trying to open Excel."
            error_handling("Error", 
                           f"{message} at DB_reader()", 
                           e,
                           is_critical = True, 
                           user_message = message)
            raise ExcelError(message) from e
        logging.info(f"Opening '{path_itdb}','{path_ctrl}'")
        itdb_sheet, itdb =  safe_book_opener(app, 
                                             path_itdb,
                                             "Item Database")
        sht, ctrl = safe_book_opener(app, 
                                    path_ctrl,
                                    "Control Doc")
        logging.info(f"Process stage completed: {process_successful}")

        '''
        -------------------------------------------------------------------------------------------------------------------------
        The below runs the sheet_copier() function, which copies the "Item Database" sheet, from the Item Database spreadsheet 
        to the Control Document. Error handling is contained within the function itself.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        if sheet_copier(itdb_sheet, 
                        path_ctrl,
                        ctrl):
            logging.info("Item database successfully copied")
        
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The following code sets the variable Formula1 to the value returned by DB_reader(), creates a selection equal to the
        length of the item database (number of rows).
        -------------------------------------------------------------------------------------------------------------------------
        '''
        Formula1= DB_reader(path_itdb, 
                            "Item Description") #Formula for drop down to be applied to
            
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The open instance of the "Control Doc" sheet is used in the drop_down() functions. They create drop down menus in the 
        'Control Document' with values from the 'Item Database'. Two separate ranges have to be given to account for the gap in 
        the 'Control Document' for extras. Again if an error occurs, the except block will catch it and display a pop up 
        notification informing the user. For testing purposes the error is printed to the terminal. In the block for the the 
        second range, the control document is saved and closed.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        try:
            #Range 1
            drop_down(sht, 
                      Formula1, 
                      "C6:C32")
            #Range 2 (Extras)
            drop_down(sht, 
                      Formula1, 
                      "C40:C50")
            ctrl.save()
            ctrl.close()
            process_successful = "Formula1"
            logging.info(f"Process stage completed: {process_successful}")

        except Exception as e:
            message = f"An error has occurred, drop down menus not created."
            error_handling("Error", 
                           message, 
                           is_critical = True, 
                           user_message = message,
                           )
            raise ExcelError(message) from e
        '''
        -------------------------------------------------------------------------------------------------------------------------
        The below code runs a clean up using the clean_up() function (defined in utilities), to remove any left over instances of
        Excel that may still be running in the background.
        -------------------------------------------------------------------------------------------------------------------------
        '''
        try:
            logging.info("Running clean up")
            if hasattr(app, 
                       'visible') and app.visible == False:
                clean_up()
        except Exception as e:
            message = f"An error has occurred when trying to close background Excel sheets. \nPlease check task manager and close any unused background sheets."
            error_handling("Error", 
                           message, 
                           e,
                           user_message = message)
    '''
    -------------------------------------------------------------------------------------------------------------------------
    The below code runs the main() and as previously mentioned allows for critical errors to stop the program and notify the
    user.
    -------------------------------------------------------------------------------------------------------------------------
    '''
    try:
        main()
        pop_up("Template Control Document Successfully Updated! ")
        logging.info("Run end ================================================================================================================================================================================================\n\n\n")
    except FileNotFoundError as e:
        pass
    except ValueError as e:
        pass
    except KeyError as e:
        pass
    except WorkBookOpenError as e:
        pass
    except Exception as e:
        message = f"Critical Error!\nProcess has been cancelled!"
        error_handling("Error", 
                       "main()", 
                       e,
                       is_critical = True, 
                       user_message = message)
        raise Exception(message) from e
    