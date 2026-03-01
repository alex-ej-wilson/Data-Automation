#Utilities

# Imports
import os.path
from tkinter import Tk, Label, Button, FLAT, LEFT, RIGHT #deafult in python
import subprocess
import xlwings as xw
import json
from concurrent.futures import ThreadPoolExecutor
import logging 
from logging.handlers import RotatingFileHandler
import sys

# Custom Error Classes ------------------------------------------------------------------
'''
Custom Error Classes allow certain errors to be flagged and then raised at the end if 
necessary
'''
class FileGenerationError(Exception):
    """
    Used for errors regarding file generation
    """
    pass
class ExcelError(Exception):
    """
    Used for errors regarding general Excel operations
    """
    pass
class SheetError(Exception):
    """
    Used for Excel sheet related errors
    """
    pass
class ConfigError(Exception):

    """
    Used for config.json related errors
    """
    pass
class CharacterError(Exception):
    
    """
    Used for special character related errors
    """
    pass
class WorkBookOpenError(Exception):
    
    """
    Used if an open workbook creates an error
    """
    pass
class BlankError(Exception):
    
    """
    Used if a blank cell creates an error
    """
    pass

# Functions -----------------------------------------------------------------------------
# =======================================================================================
# Logs ----------------------------------------------------------------------------------
def logger(__file__):
    '''
    Creates the log file and configures the logger

    Usage: logger(__file__)
    '''
    name = os.path.basename(__file__)
    name = name.replace('.py','')
    log_directory = full_path_maker(f"logs\\{name}_logs")
    log_path = f"{log_directory}\\{name}.log"
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
    
    
    handler = RotatingFileHandler(
        log_path,
        maxBytes=10*1024*1024,
        backupCount=5
    )

    handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    logging.getLogger().addHandler(handler)
    logging.getLogger().setLevel(logging.INFO)
    logging.info(f"{name}: Run start ================================================================================================================================================================================================\n")
# File Paths ----------------------------------------------------------------------------
def truncater(file,root):
    '''
    Removes the root_path from the file_path to give the shortened file address

    Usage: truncator(file_path, root_path)

    Returns: a shortened file path (i.e. without the root)
    '''
    f = file.replace(root,"")
    return f
def full_path_maker(given_path):
    '''
    Takes a file name, assuming its in the same directory as the script and produces the
    full path
    '''
    script_dir = os.path.dirname(os.path.abspath(__file__)) 
    full_path = os.path.join(script_dir, 
                             given_path)
    if os.path.exists(full_path):
        return full_path
    else:
        message = f"Full path made: '{full_path}' does not exist. "
        error_handling("Error", 
                       f"full_path_maker(): {message}",
                       "Path does not exist", 
                       is_critical = True)
        raise FileNotFoundError
def get_excel_file_path(path):
            try:
                logging.info("Reading arguments from Excel sheet")
                path = sys.argv[1]  #Excel Book has passed its file path to system arguments
            except IndexError as e: 
                message = f"Error finding Excel file path from system arguments. Using test file path: {path}"
                error_handling("Error", 
                               message, 
                               e,
                               is_critical = False, 
                               user_message = message)
            if not os.path.exists(path):
                message = f"File path invalid: '{path}'"
                error_handling("Error", 
                               message, 
                               e,
                               is_critical = True, 
                               user_message = message)
                raise FileNotFoundError(message) from e
            return path

# Openers/Closers -----------------------------------------------------------------------
def json_reader(config_name):
    '''
    Reads .json files, and outputs them as a dictionary

    Usage: json_reader(config_file_name.json)
    
    Returns: .json file contents as a dictionary
    '''
    config_path = full_path_maker(config_name)
    try:
        logging.info(f"Reading config.json: {config_path}")
        with open(config_path,"r") as f:
            config = json.load(f)
        return config, config_path
    except FileNotFoundError as e:
        message = f"Config file not found at {config_path}"
        error_handling("Error", 
                       message, 
                       e, 
                       is_critical = True, 
                       user_message = message)
    except KeyError as e:
        message = f"An error has occurred with the 'config.json' file"
        error_handling("Error", 
                       message, 
                       e, 
                       is_critical = True, 
                       user_message = message)
    except Exception as e:
        message = f"An unexpected error has occurred with the 'config.json' file, please ensure all formatting is correct."
        error_handling("Error", 
                       message, 
                       e, 
                       is_critical = True, 
                       user_message = message)
def clean_up():
    '''
    Cleans up any remaining Excel instances.

    Usage: clean_up()
    '''
    try:
        book = app.books.add()
        sheet = book.sheets[0]
        sheet.range('A1').value = 'Test'
        app.quit()
    except NameError as e:
        print("Nothing to clean up!")
        logging.info(f"Nothing to clean up: {e}")

# GUI -----------------------------------------------------------------------------------  
thread_pool = ThreadPoolExecutor(max_workers=3) 
def create_button(window, 
                  text_input, 
                  command_input):
    '''
    Creates tkinter button with specific preferences.

    Usage: clean_up(window, message, command)
    '''
    return Button(window, 
                  text = text_input,
                          font = ('Arial','9'), 
                          command = command_input, 
                          height= 1, relief=FLAT, 
                          bg = ("gainsboro"), cursor="hand2",
                          activeforeground="deep sky blue", 
                          activebackground="light sky blue", 
                          highlightcolor="deep sky blue", 
                          highlightthickness=3,bd=0)
def pop_up(message, 
           close = True, 
           file = None, 
           option = False, 
           t = 7000, 
           async_mode = True):
    def create_window():
        window = Tk()
        window.title("Information")
        window.geometry("290x140")
        window.resizable(False, False)
        window.eval(f"tk::PlaceWindow {window.winfo_toplevel()} centre")
        def file_opener(file_):
            '''
            Opens given file

            Usage: file_opener(file_path)
            '''
            subprocess.Popen(f'explorer /select,{file_}')
            window.destroy()
        flag_dict = {"flag" : False}
        def flag():
                flag_dict["flag"] = True
                window.destroy()
        label = Label(window, 
                      font = ('Arial' ,'10'), 
                      text = message, 
                      wraplength = 250, 
                      width =40, 
                      anchor ="center",
                      justify = "center")
        label.pack(pady=10,
                   padx= 7,
                   side = "top")
        if close:
                close_button = create_button(window, 
                                             "Close", 
                                             window.destroy)
                close_button.pack(pady=0)
        if file:
            file_button = create_button(window, 
                                        "Open file location", 
                                        lambda: file_opener(file))
            file_button.pack(pady = 7, 
                             side = LEFT, 
                             padx = 30)
            close_button.pack(pady = 7,
                              side = RIGHT, 
                              padx = 30)
        if option:
            yes_button = create_button(window,
                                       "Yes", 
                                       flag)
            yes_button.pack(pady = 7, 
                            side = LEFT, 
                            padx = 30)
            no_button = create_button(window, 
                                      "No", 
                                      window.destroy)
            no_button.pack(pady = 7,
                           side = RIGHT,
                           padx = 30)
        if t > 0:
            window.after(t, 
                         window.destroy)
        window.mainloop()
        return flag_dict["flag"]
    if async_mode:
        # Use threading to run the pop-up asynchronously
        thread_pool.submit(create_window)
    else:
        # Run the pop-up synchronously and return the flag
        return create_window()

# Errors and Validation -----------------------------------------------------------------     
def error_handling(error_type, 
                   context, error = None, 
                   is_critical = False, 
                   user_message = None, 
                   file = None,
                   time = 0):
    '''
    A centralised error handler with optional pop up. Will log errors and warnings, and option to raise as critical.

    Usage: error_handing(error_type, context, error, is_critical = True/False, user_message = message)
    '''
    if error:
        if error_type == "Warning":
            logging.warning(f"{context}: {type(error).__name__}:{error}",exc_info=True)
            if user_message:
                if file:
                    pop_up(f" ⚠️ {user_message}", 
                           file,
                           t = time)
                else:
                    pop_up(f" ⚠️ {user_message}",
                           t = time)
        if error_type == "Error":
            logging.error(f"{context}: {type(error).__name__}:{error}",exc_info=True)
            if user_message:
                if file:
                    pop_up(f" ❌ {user_message}", 
                           file,
                           t = time)
                else:
                    pop_up(f" ❌ {user_message}",
                           t = time)
        if is_critical:
            raise error
    else:
        if error_type == "Warning":
            logging.warning(f"{context}")
            if user_message:
                if file:
                    pop_up(f" ⚠️ {user_message}", 
                        file,
                        t = time)
                else:
                    pop_up(f" ⚠️ {user_message}",
                           t = time) 
        if error_type == "Error":
            logging.error(f"{context}")
            if user_message:
                if file:
                    pop_up(f" ❌ {user_message}", 
                           file,
                           t = time)
                else:
                    pop_up(f" ❌ {user_message}",
                           t = time)
def safe_sheet_opener(book_,
                      sheet_):
    '''
    Takes a workbook, sheet name and file path, and returns that sheet as a sheet object.
    The function takes care of error handling and logging to catch more errors without 
    repeated code.

    Usage: safe_sheet_opener(workbook, sheet name, file path)
    '''
    validate_sheet_name(book_,
                        sheet_)
    logging.info(f"Opening {sheet_} from {book_}")
    sheet_object_ = book_.sheets[sheet_]
    return sheet_object_
def safe_book_opener(application, 
                     path, 
                     sheet = None):
    '''
    Takes an application (variable initialised in the script), file path (and optional sheet name)
    to return that workbook as an opened workbook object. If a sheet name is given it will return 
    the sheet and workbook as a tuple.

    The function takes care of error handling and logging to catch more errors without 
    repeated code.

    Usage: safe_book_opener(app variable, file path, sheet name)
    '''
    try:
        logging.info(f"Opening book at: '{path}'")
        book = application.books.open(path)
        if sheet:
            # Creating sheet object to return        
            sheet_object = safe_sheet_opener(book,
                                             sheet)
            return sheet_object, book
        else:
            return book
    except FileNotFoundError as e:
        error_handling("Error", 
                       "safe_book_opener(): file not found at '{path}'. Error: ", 
                       e,
                       is_critical = True, 
                       user_message = "File not found!")
    except KeyError as e:
        message = "Application error!\Incorrect sheet name used, check config.json"
        error_handling("Error", 
                       "Application not found. Error: ", 
                       e, 
                       is_critical = True, 
                       user_message = message)
        raise ExcelError(message)
    except Exception as e:
        error_handling("Error", 
                       "Unexpected Error: ", 
                       e,
                       is_critical = True, 
                       user_message = "An unexpected error has occurred!")
def safe_renamer(old_name,
                 new_name):
            logging.info(f"Renaming control doc from {old_name} to {new_name}")
            if os.path.exists(new_name):
                pop_up(f"Control document already named '{new_name}'")
            else:
                os.rename(old_name,
                          new_name)
def validate_sheet_name(workbook, 
                        sheet_name):
    '''
    Checks to see if a sheet exists in a specified workbook, if it does not a SheetError is raised

    Usage: validate_sheet_name(workbook, sheet_name)
    '''
    if sheet_name not in [sheet.name for sheet in workbook.sheets]:
        raise SheetError(f"Sheet '{sheet_name}' does not exist in '{workbook}' Workbook.")

if __name__ == "__main__":
        pop_up("Pop up Test", 
               t = 1000)
        print("DONE")


