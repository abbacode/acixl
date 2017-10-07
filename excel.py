import xlwings as xw

# Workbook details
WORKBOOK_NAME = 'runsheet.xlsm'
wb = xw.Book(WORKBOOK_NAME)
ws_apic = wb.sheets['APIC']

# APIC authentication cell locations
APIC = ws_apic.range('$B$5').value
USER = ws_apic.range('$B$6').value
PWORD = ws_apic.range('$B$7').value

# Control panel cell locations
AUTH_STATUS_1 = '$B$3'
AUTH_STATUS_2 = '$B$4'
SCRIPT_STATUS_1 = '$C$5'
SCRIPT_STATUS_2 = '$C$6'
CONSOLE_CELL = '$E$3'

# Font Colors
COLOR_BLACK = 1
COLOR_BLUE = 5

# Background colors for cells
COLOR_DEFAULT = (255, 255, 255, 255)  # white
COLOR_FAILED = (250, 128, 144)  # Red
COLOR_PASS = (144, 238, 144)  # Green
COLOR_IGNORED = (255, 207, 80)  # Amber

HTTP_STATUS_CODES = {200: {'msg1': '200 - Success (OK)',
                           'msg2': 'Authentication Successful',
                           'color': COLOR_PASS},
                     400: {'msg1': '400 - Bad request',
                           'msg2': 'Bad URL or payload',
                           'color': COLOR_FAILED},
                     401: {'msg1': '401 - Unauthorised',
                           'msg2': 'Bad credentials',
                           'color': COLOR_FAILED},
                     403: {'msg1': '403 - Forbidden',
                           'msg2': 'Server refusing to handle request',
                           'color': COLOR_FAILED},
                     404: {'msg1': '404 - Not found',
                           'msg2': 'Post to page that does not exist',
                           'color': COLOR_FAILED},
                     999: {'msg1': '999 - Unknown error occured',
                           'msg2': 'Check IP/connectivity',
                           'color': COLOR_FAILED}}

# Control panel cells to reset when button is pushed
CP_CELLS_TO_RESET = (AUTH_STATUS_1, AUTH_STATUS_2,
                     SCRIPT_STATUS_1, SCRIPT_STATUS_2)

TABLE_TENANT_POLICIES = ["TABLE_TENANT", 'TABLE_VRF', 'TABLE_ANP',
                         'TABLE_BD', 'TABLE_EPG', 'TABLE_BD_SUBNET']

TABLE_FABRIC_ACCESS_POLICIES = ["TABLE_CDP"]

TABLES = {'Tenant_Policies': TABLE_TENANT_POLICIES,
          'Fabric_Access_Policies': TABLE_FABRIC_ACCESS_POLICIES}

# delete table rows that have missing mandatory keys
def remove_invalid_rows(table, mandatory_keys):
    # make a copy of the table
    table_copy = dict(table)
    for row in sorted(table_copy):
        for key in mandatory_keys:
            if table_copy[row].get(key) == None:
                status_cell = table_copy[row]['status_cell']
                if table.get(row):
                    del table[row]
                #console_msg = ''
                #console_msg += 'row: {} missing: {}'.format(row,key)
                #update_console(console_msg)
                update_cell(cell=status_cell,
                            value='Row ignored - missing field',
                            bg_color=COLOR_IGNORED)
    return table


# update empty, non-mandatory rows with default values
def apply_default_values(table, default_values):
    for row in sorted(table):
        for d in default_values:
            #console_msg+='\nprocessing: {}'.format(d)
            #update_console(msg=console_msg)
            value = table[row].get(d)
            #console_msg += '\n  --current value: {}'.format(value)
            if value == None:
                default_value = default_values[d]
                table[row][d] = default_value
                #console_msg+='\n     Value changed to: {}'.format(default_value)
                #update_console(msg=console_msg)

        # Convert remaining rows that have value of None to empty string: ''
        for k,v in sorted(table[row].items()):
            if v == None:
                #console_msg += '\nempty value for key: {}, changed to string'.format(k)
                #update_console(msg=console_msg)
                table[row][k] = ''
    return table



def get_table(table_name=None, mandatory_keys=None, default_values=None):
    t = xw.Range(table_name)
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


# update the value or color of a cell
def update_cell(cell='', value='', bg_color=COLOR_DEFAULT, font_color=COLOR_BLACK):
    cell = xw.Range(cell)
    cell.value = value
    cell.api.Font.ColorIndex = font_color
    cell.color = bg_color


# update the status_code column value based on status_code
def update_status(cell, status_code):
    #get the msg and color based on the http status code returned
    status = HTTP_STATUS_CODES.get(status_code)['msg1']
    bg_color = HTTP_STATUS_CODES.get(status_code)['color']

    update_cell(cell=cell,value=status,bg_color=bg_color)


# return a list of status codes for all rows in a table
def get_table_status_code(table_name):
    table = xw.Range(table_name)
    status_code = [row[0] for row in table.value[2:]]
    return status_code

def show_auth_attempt_msg():
    update_cell(cell=AUTH_STATUS_1,value='Attempting authentication...')
    update_cell(cell=AUTH_STATUS_2,value='Do not push policies unless this works')


# update the control plane status field whenever a script is executed
def update_cp_status(table_name, script_msg='N/A'):
    table_status = get_table_status_code(table_name)

    # update the cell to reference what action was executed
    update_cell(cell=SCRIPT_STATUS_1, value=script_msg, font_color=1)

    # update the cell to indicate how successful it was
    if all(status == 200 for status in table_status):
        update_cell(cell=SCRIPT_STATUS_2,
                    value='All entries pushed',
                    bg_color=COLOR_DEFAULT,
                    font_color=1)
    elif any(status == 200 for status in table_status):
        update_cell(cell=SCRIPT_STATUS_2,
                    value='Partial entries pushed',
                    bg_color=COLOR_DEFAULT,
                    font_color=1)
    else:
        update_cell(cell=SCRIPT_STATUS_2,
                    value='All entries failed',
                    bg_color=COLOR_DEFAULT,
                    font_color=1)

# update authentication status field based on status code
def update_auth_status(status_code):
    show_auth_attempt_msg()
    if HTTP_STATUS_CODES.get(status_code):
        msg_1 = HTTP_STATUS_CODES.get(status_code)['msg1']
        msg_2 = HTTP_STATUS_CODES.get(status_code)['msg2']
        bg_color = HTTP_STATUS_CODES.get(status_code)['color']
        update_cell(cell=AUTH_STATUS_1, value=msg_1, bg_color=bg_color)
        update_cell(cell=AUTH_STATUS_2, value=msg_2, bg_color=bg_color)


# change the action field for a table
def set_table_action(table_name, action):
    table = xw.Range(table_name)
    # skip the first two rows (headers)
    for no, row in enumerate(table.value[2:]):
        cell = table[no + 2, 1].address
        color = table[no + 2, 1].color
        xw.Range(cell).value = action
        xw.Range(cell).color = color


def set_all_table_action(action):
    current_worksheet = xw.sheets.active.name
    console_msg = 'Change action for all tables in worksheet: {}'.format(current_worksheet)
    update_console(console_msg)
    for table in TABLES[current_worksheet]:
        console_msg += '\n  -{} action changed to --> \'{}\''.format(table, action)
        update_console(console_msg)
        set_table_action(table, action)


def update_console(msg):
    update_cell(cell=CONSOLE_CELL, value=msg, font_color=COLOR_BLUE)


def show_console_payload(row, table_name, uri, payload):
    msg = 'Reading row: {}, from table: {}'.format(row + 1, table_name)
    msg += '\n\nPayload: {}'.format(payload)
    msg += '\n\nPosting to URI: {}'.format(uri)
    update_console(msg=msg)


def update_console_cmd_not_found(cmd_name):
    msg = 'Error, could not find: {}'.format(cmd_name)
    update_console(msg=msg)


# reset the status fields in the control panel
def reset_table_control_panel():
    for cell in CP_CELLS_TO_RESET:
        update_cell(cell=cell, value='')


# reset the status fields in the control panel
def reset_console():
    update_cell(cell=CONSOLE_CELL, value='')


# reset the status_code column for all tables
def reset_table_status():
    try:
        current_worksheet = xw.sheets.active.name
        console_msg = 'Clear status code for worksheet: {}'.format(current_worksheet)
        update_console(console_msg)
        for table in TABLES[current_worksheet]:
            console_msg += '\n  -Clearing status code for table: {}'.format(table)
            update_console(console_msg)
            t = xw.Range(table)
            for no, row in enumerate(t.value[2:]):
                cell = t[no + 2, 0].address
                xw.Range(cell).value = ''
                xw.Range(cell).color = COLOR_DEFAULT
    except Exception as e:
        console_msg = '  \n-Error clearing status codes'
        update_console(console_msg)


def reset_all_status():
    reset_table_control_panel()
    reset_table_status()
    # reset_console()
