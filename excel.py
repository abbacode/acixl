import xlwings as xw
from xlwings.constants import DeleteShiftDirection

# Workbook details
WORKBOOK_NAME = 'runsheet.xlsm'
wb = xw.Book(WORKBOOK_NAME)

# Define hidden internal worksheet names
ws_commands = wb.sheets['_commands']
ws_tables = wb.sheets['_tables']

# Get the authentication details from the spreadsheet
APIC = wb.sheets['Test_Authentication'].range('$B$3').value
USER = wb.sheets['Test_Authentication'].range('$B$4').value
PWORD = wb.sheets['Test_Authentication'].range('$B$5').value

# Define special cell locations
CONSOLE_CELL = '$E$3'

# Font Colors
COLOR_BLACK = 1
COLOR_RED = 2
COLOR_BLUE = 5

# Background colors for cells
COLOR_DEFAULT = (255, 255, 255, 255)  # white
COLOR_FAILED = (250, 128, 144)  # Red
COLOR_PASS = (144, 238, 144)  # Green
COLOR_IGNORED = (255, 207, 80)  # Amber

# HTTP status codes, meanings and formatting
HTTP_STATUS_CODES = {200: {'msg1': '200',
                           'msg2': 'Success',
                           'color': COLOR_PASS},
                     400: {'msg1': '400',
                           'msg2': 'Bad request - incorrect URL or payload',
                           'color': COLOR_FAILED},
                     401: {'msg1': '401',
                           'msg2': 'Unauthorised - wrong credentials',
                           'color': COLOR_FAILED},
                     403: {'msg1': '403',
                           'msg2': 'Forbidden - server refusing to handle request',
                           'color': COLOR_FAILED},
                     404: {'msg1': '404',
                           'msg2': 'Not found - Post to page that does not exist',
                           'color': COLOR_FAILED},
                     999: {'msg1': '999 - Unknown error occured',
                           'msg2': 'Check IP/connectivity',
                           'color': COLOR_FAILED}}


def remove_invalid_rows(table, mandatory_keys):
    """
    Iterate through a table to see whether any of the mandatory keys
    (defined in launcher.json) are missing. If so, invalidate the rows
    and update the status_cell in excel to inform the user that the row
    will not be processed as part of the post operation.

    Args:
        table(dict): the table as called via the get_table() function
        mandatory_keys (list): as extracted from launcher.json

    Returns:
        table(dict): revised dictionary minus the invalid row entries

    """
    table_copy = dict(table)
    for row in sorted(table_copy):
        for key in mandatory_keys:
            if not table_copy[row].get(key):
                status_cell = table_copy[row]['status_cell']
                if table.get(row):
                    del table[row]
                update_cell(cell=status_cell,
                            value='Row ignored - missing field',
                            bg_color=COLOR_IGNORED)
    return table


def apply_default_values(table, default_values):
    """
    Iterate through a table and use the default values as defined in
    launcher.json for non-mandatory columns. This is required otherwise
    the payload would not be constructed correctly and the post would fail.

    For the remaining non-mandatory columns that have no default value
    defined, replace the cell content with an empty string, otherwise
    'None' would be used, breaking the payload structure.

    Args:
        table(dict): the table as called via the get_table() function
        default_values (dict): as extracted from launcher.json

    Returns:
        table(dict): revised dictionary with updated default values

    """
    for row in sorted(table):
        for d in default_values:
            current_value = table[row].get(d)
            if not current_value:
                new_value = default_values[d]
                table[row][d] = new_value

        # Convert remaining rows that have value of None to empty string: ''
        for k,v in sorted(table[row].items()):
            if not v:
                table[row][k] = ''
    return table


def get_table(table_name=None, mandatory_keys=None, default_values=None):
    """
    Read the content of a table from the active worksheet in excel.
    The table must have a dynamic name assigned to it. This function
    will create a dictionary with the row as the primary key, columns
    as the sub-keys which are mapped to the cell content.

    Args:
        table_name(str): The name of the table in excel.
        mandatory_keys (list): table mandatory keys, from launcher.json
        default_values (dict): table default values, from launcher.json

    Returns:
        table(dict): content from table in excel that has removed
        invalid rows and has default values applied to cells

    """
    t = xw.Range(table_name).options(numbers=int)
    header = xw.Range(table_name).value[0]
    table = {}
    # skip the first two rows (headers)
    for row, col in enumerate(t.value[2:]):
        dictionary = dict(zip(header, col))
        table[row] = {}
        table[row] = dictionary
        table[row]['status_cell'] = t[(row + 2, 0)].address

    # remove invalid rows
    table = remove_invalid_rows(table, mandatory_keys=mandatory_keys)

    # if a non-mandatory cell has no value, then use default values
    table = apply_default_values(table, default_values=default_values)

    return table


def get_table_list():
    """
    Read the 'TABLES' table found under the hidden '_tables' worksheet.
    This table is populated from launcher.json file and is used to
    determine the table names defined under each worksheet

    Returns:
        table(dict): k,v, k=worksheet_name and v=list of tables

    """
    tables = ws_tables.range('TABLES').value
    table_list= {}
    for t in tables:
        worksheet_name,table_name = t[0], t[1]
        if not table_list.get(worksheet_name):
            table_list[worksheet_name] = []
        table_list[worksheet_name].append(table_name)
    return table_list


def update_cell(cell='', value='', bg_color=COLOR_DEFAULT, font_color=COLOR_BLACK):
    """
    Updates a cell location in the active worksheet with a text value

    Args:
        cell(str): cell, i.e. A1, or (1,1)
        value(str): the text to place in the cell
        bg_color(int, optional): background color of cell
        font_color(int, optional): font color for text in cell

    """
    cell = xw.Range(cell)
    cell.value = value
    cell.api.Font.ColorIndex = font_color
    cell.color = bg_color


def update_status(cell, status_code):
    """
    Updates a cell with pre-formatted colors and msg based on the
    status_code received when attempting to post the payload

    Args:
        cell(str): cell, i.e. A1, or (1,1)
        status_code(int): status code number returned from REST post command

    """
    status = HTTP_STATUS_CODES.get(status_code)['msg1']
    bg_color = HTTP_STATUS_CODES.get(status_code)['color']
    update_cell(cell=cell, value=status, bg_color=bg_color)


def get_status_codes_from_table(table_name):
    """
   Get all the status codes from the status_code column for a table

    Args:
        table_name(str): The name of the table in excel. The table must have
        a label assigned to it, otherwise the function won't find it.

    Returns:
        status_codes (list): return the status codes from a table

    """
    table = xw.Range(table_name).options(numbers=int)
    status_codes = [str(row[0]) for row in table.value[2:]]
    return sorted(status_codes)

def get_failed_rows_from_table(table_name):
    """
    Check the status_code column for a specific table and grab a
    all list of rows which did not push successfully

    Args:
        table_name(str): The name of the table in excel. The table must have
        a label assigned to it, otherwise the function won't find it.

    Returns:
        failed_rows (list): return the failed rows from the table

    """
    table = xw.Range(table_name).options(numbers=int)
    failed_rows = [row_no+1 for row_no, row in enumerate(table.value[2:])
                   if row[0] != 200]
    return sorted(failed_rows)


def show_push_report_status(table_name, action_msg):
    """
    Updates the console in the control panel with a list of rows
    which did not execute successfully as part of the push. This function
    will not do anything if there were no failed rows.

    Args:
        table_name(str): Name of the table which is retrieved from
        launcher.json file

        action_msg(str, optional): msg to be shown in the 'last
        action performed' cell,  passed from launcher.json file

    """
    # get the inital console msg
    console_msg = get_status_results(table_name, action_msg)

    # get a list of failed rows
    failed_rows = get_failed_rows_from_table(table_name)

    # add failed rows output to the console_msg (if there are any)
    if failed_rows:
        console_msg += '\n\nThe following rows from table {} experienced ' \
                      'a problem'.format(table_name)
        for row in failed_rows:
            console_msg += '\n -- Row: {} did not post properly'.format(row)
    update_console(msg=console_msg)


def get_status_results(table_name, action_msg=''):
    """
    Updates the control panel script status each time an action is performed.
    It retrieves the status column for a given table and then displays
    the result, i.e. complete success, partial success or compelte failure.

    Args:
        table_name(str): Name of the table which is retrieved from
        launcher.json file

        action_msg(str, optional): msg to be shown in the 'last
        action performed' cell,  passed from launcher.json file

    """
    # get the output of the status_column in list format for a specified table
    table_status = get_status_codes_from_table(table_name)

    console_msg = 'Last action performed: {}'.format(action_msg)

    # update the cp 'script status' cell to success status:
    if all(status == '200' for status in table_status):
        console_msg +='\n  -- Action status: all entries posted to APIC'
    elif any(status == '200' for status in table_status):
        console_msg += '\n  -- Action status: partial entries pushed'
    else:
        console_msg += '\n  -- Action status: all entries failed'
    return console_msg

def show_cp_authentication_attempt_msg():
    """
    Show the authentication attemp message
    """
    update_console('Attempting authentication...')


# update authentication status field based on status code
def update_cp_authentication_response(status_code):
    """
    Updates the authentication success status in the control panel
     based on the status code received from the attempt.

    Args:
        status_code(int): status code response received from auth attempt

    """
    show_cp_authentication_attempt_msg()
    if HTTP_STATUS_CODES.get(status_code):
        msg_1 = HTTP_STATUS_CODES.get(status_code)['msg1']
        msg_2 = HTTP_STATUS_CODES.get(status_code)['msg2']
        console_msg = 'Authentication response from APIC'
        console_msg += '\n  - Status code: {}'.format(msg_1)
        console_msg += '\n  - Status explanation: {}'.format(msg_2)
        update_console(msg=console_msg)


def set_table_action(table_name, action):
    """
    Update the value of the 'action' column for a specific table
    within the active worksheet

    Args:
        table_name(str): The name of the table in excel. The table must have
         a label assigned to it, otherwise the function won't find it.

        action(str): Text value that is added to the action column

    """
    table = xw.Range(table_name)
    # skip the first two rows (headers)
    for no, row in enumerate(table.value[2:]):
        cell = table[no + 2, 1].address
        color = table[no + 2, 1].color
        xw.Range(cell).value = action
        xw.Range(cell).color = color


def set_all_table_action(action):
    """
    Update the value of the 'action' column for all tables in the active worksheet

    Args:
        action(str): Text value that is added to the action column

    """
    current_worksheet = xw.sheets.active.name
    console_msg = 'Change action for all tables in worksheet: {}'.format(current_worksheet)
    update_console(console_msg)

    TABLES = get_table_list()
    for table in TABLES[current_worksheet]:
        console_msg += '\n  -{} action changed to --> \'{}\''.format(table, action)
        update_console(msg=console_msg)
        set_table_action(table, action)


def update_console(msg):
    """
    Updates the console cell in the control panel of the active worksheet

    Args:
        msg(str): Message that will be generated and posted to the console cell

    """
    update_cell(cell=CONSOLE_CELL, value=msg, font_color=COLOR_BLUE)


def show_console_payload(row, table_name, uri, payload):
    """
    Update the console to show how the payload has been constructed,
    and the URI it will be posted to.

    Args:
        row(int): Row of the table used to construct the payload
        table_name(str): Name of the table
        uri(str): Full URI that the payload is posted to
        payload(str): JSON format of the fully formed payload

    """
    console_msg = 'Reading row: {}, from table: {}'.format(row + 1, table_name)
    console_msg += '\n\nPayload: {}'.format(payload)
    console_msg += '\n\nPosting to URI: {}'.format(uri)
    update_console(msg=console_msg)


def show_console_launcher_error(launcher_fname=None):
    """
    Update the console with an error that the launcher.json
    could not be found.
    """
    msg = 'Error, could not find: {}'.format(launcher_fname)
    update_console(msg=msg)


def reset_cp_console():
    """
    Resets the console cell of the active worksheet
    """
    update_cell(cell=CONSOLE_CELL, value='')


def reset_table_status():
    """
    Cycle through all the tables of the active worksheet that is currently open,
    and clear the text of the status_code column (first column)
    """
    try:
        current_worksheet = xw.sheets.active.name
        console_msg = 'Clear status code for worksheet: {}'.format(current_worksheet)
        update_console(console_msg)
        TABLES = get_table_list()
        for table in TABLES[current_worksheet]:
            console_msg += '\n  -Clearing status code for table: {}'.format(table)
            update_console(console_msg)
            t = xw.Range(table)
            for no, row in enumerate(t.value[2:]):
                cell = t[no + 2, 0].address
                xw.Range(cell).value = ''
                xw.Range(cell).color = COLOR_DEFAULT
    except Exception as e:
        update_console('\n  -Error clearing status codes')


def reset_all_status():
    """
    Executes the various status clearing functions
    """
    reset_table_status()
    reset_cp_console()

def update_excel_data(commands, tables):
    """
    Updates the table in the hidden '_command' worksheet with the key values retrieved
    from the launcher.json. These keys are presented as commands in the drop
    down list located in the 'post to apic' section of the control panel.

    Args:
        cmd_list(list): List of keys taken from launcher.json file

    """
    worksheet_names = tables[0]
    worksheet_tables = tables[1]

    # add a row in case the table is empty, otherwise delete would return error
    ws_commands.range('COMMANDS').value = 'N/A'
    ws_commands.range('COMMANDS').api.Delete(DeleteShiftDirection.xlShiftUp)
    ws_commands.range('COMMANDS').options(
        transpose=True,ndim=1).value = sorted(commands)

    # add a row in case the table is empty, otherwise delete would return error
    ws_tables.range('TABLES').value = 'N/A'
    ws_tables.range('TABLES').api.Delete(DeleteShiftDirection.xlShiftUp)
    ws_tables.range('TABLES').options(
        transpose=True,ndim=1).value = [worksheet_names,worksheet_tables]

    update_console(msg='Command list refreshed from launcher.json')
    #TODO: prepend the fabric folder name to the command


def can_run_cmd_from_worksheet(cmd, launcher_data):
    # get the current active worksheet and the cmd allowed worksheet names
    current_worksheet = xw.sheets.active.name
    allowed_worksheet = launcher_data.get(cmd)['worksheet_name']

    # if they do not match, prevent the action from being performed
    if allowed_worksheet != current_worksheet:
        console_msg = 'Push configuration failed.\n'
        console_msg += '  - Cmd: \'{}\' must be pushed from the \'{}\' ' \
                       'worksheet'.format(cmd, allowed_worksheet)
        update_console(msg=console_msg)
        return False
    return True
