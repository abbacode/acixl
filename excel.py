import xlwings as xw

WORKBOOK_NAME = 'runsheet.xlsm'

wb = xw.Book(WORKBOOK_NAME)
ws_apic   = wb.sheets['APIC']

# APIC authentication cell locations
APIC  = ws_apic.range('$B$5').value
USER  = ws_apic.range('$B$6').value
PWORD = ws_apic.range('$B$7').value

# Control panel cell locations
AUTH_STATUS_1 = '$B$3'
AUTH_STATUS_2 = '$B$4'
SCRIPT_STATUS_1 = '$C$5'
SCRIPT_STATUS_2 = '$C$6'
CONSOLE_CELL = '$E$3'

# Control panel cells to reset when button is pushed
CP_CELLS_TO_RESET = (AUTH_STATUS_1, AUTH_STATUS_2, SCRIPT_STATUS_1,
                     SCRIPT_STATUS_2, CONSOLE_CELL)

# Table named ranges used in excel
TABLE_NAMES = ('TABLE_TENANT','TABLE_VRF','TABLE_ANP','TABLE_BD',
               'TABLE_EPG','TABLE_BD_SUBNET')

# Default values to use if the row has an empty, non mandatory cell
DEFAULT_VALUES = {'action': 'created',
                  'description': '',
                  'policy_enforce': 'enforced',
                  'policy_direction': 'ingress',
                  'l2_unknown_unicast': 'proxy',
                  'l3_unknown_mcast': 'flood',
                  'mdst_flooding': 'bd-flood',
                  'arp_flooding': 'false',
                  'limit_iplearn_subnet': 'no',
                  'unicast_routing': 'true',
                  'treat_as_virtual_ip': 'false',
                  'make_primary_ip': 'false',
                  'private_to_vrf': 'false',
                  'advertised_externally': 'false',
                  'shared_between_vrfs': 'false'}

# Cells that are mandatory, otherwise ignore the row
MANDATORY_VALUES = ('tn_name', 'anp_name', 'bd_name',
                    'vrf_name', 'epg_name', 'subnet')

# Font Colors
COLOR_BLACK = 1
COLOR_BLUE = 5

# Background colors for cells
COLOR_DEFAULT = (255,255,255,255) #white
COLOR_FAILED = (250,128,144)      # Red
COLOR_PASS = (144,238,144)        # Green
COLOR_IGNORED = (255,207,80)      # Amber



# remove rows that do not have mandatory values from a table
def remove_invalid_rows(table):
    temp_table = dict(table)
    #print ('temp_table: ',temp_table)
    for row in sorted(temp_table):
        #print ('processing row: ',row)
        for key in temp_table[row].keys():
            if key in MANDATORY_VALUES:
                value = temp_table[row][key]
                if not value:
                    status_cell = temp_table[row]['status_cell']
                    if table.get(row):
                        #print('mandatory cell missing, deleting row: {}'.format(row))
                        del table[row]
                    update_cell(cell=status_cell,
                                value='Row ignored - missing fields',
                                bg_color=COLOR_IGNORED)
    return table


# update table data to use default ACI values if non-mandatory empty cells are found
def apply_default_values(table):
    temp_table = dict(table)
    for row in temp_table:
        for key in temp_table[row].keys():
            if key in DEFAULT_VALUES.keys():
                value = temp_table[row][key]
                if not value:
                    table[row][key] = DEFAULT_VALUES[key]
    return table



#------ new and ready to go

def get_table(table_name):
    t = xw.Range(table_name)
    header = xw.Range(table_name).value[0]
    table = {}
    #skip the first two rows (headers)
    for row,col in enumerate(t.value[2:]):
        dictionary = dict(zip(header, col))
        table[row] = {}
        table[row] = dictionary
        table[row]['status_cell'] = t[(row+2,0)].address
    # remove invalid rows
    table = remove_invalid_rows(table)
    # if a non-mandatory cell has no value, then use default values
    table = apply_default_values(table)
    return table

# update the value or color of a cell
def update_cell(cell='', value='', bg_color=COLOR_DEFAULT, font_color=COLOR_BLACK):
    cell = xw.Range(cell)
    cell.value = value
    cell.api.Font.ColorIndex = font_color
    cell.color = bg_color

# update the status_code column value based on status_code
def update_cell_status(cell, status_code):
    cell = xw.Range(cell)
    cell.value = status_code
    cell.color = get_status_color(status_code)

# return a list of status codes for all rows in a table
def get_status_codes_from_table(table_name):
    table = xw.Range(table_name)
    status_code = [row[0] for row in table.value[2:]]
    return status_code


# update the control plane status field whenever a script is executed
def update_cp_status(table_name, script_msg='N/A'):
    table_status = get_status_codes_from_table(table_name)

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

def get_status_color(status_code):
    COLORS = {200: COLOR_PASS,
              'Row ignored - missing fields': COLOR_IGNORED}
    color = COLORS.get(status_code, COLOR_FAILED)
    return color


def show_auth_attempt_msg():
    update_cell(cell=AUTH_STATUS_1,
                value='Attempting authentication...')
    update_cell(cell=AUTH_STATUS_2,
                value='Do not push policies unless this works')

# update authentication status field based on status code
def update_auth_status(status_code):
    show_auth_attempt_msg()
    STATUS = {200: {'msg1':'200 - Authentication successful',
                    'msg2':'Received token from APIC',
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

    if STATUS.get(status_code):
        msg_1 = STATUS.get(status_code)['msg1']
        msg_2 = STATUS.get(status_code)['msg2']
        status_color = STATUS.get(status_code)['color']
        update_cell(cell=AUTH_STATUS_1, value=msg_1, bg_color=status_color)
        update_cell(cell=AUTH_STATUS_2, value=msg_2, bg_color=status_color)


# change the action field for a table
def set_table_action(table_name, action):
    table = xw.Range(table_name)
    # skip the first two rows (headers)
    for no,row in enumerate(table.value[2:]):
        cell = table[no+2,1].address
        color = table[no+2,1].color
        xw.Range(cell).value = action
        xw.Range(cell).color = color


def set_all_table_action(action):
    for table in TABLE_NAMES:
        set_table_action(table, action)

def update_console(row, table_name, uri, payload):
    console_msg = 'Reading row: {}, from table: {}'.format(row+1,table_name)
    console_msg += '\n\nPayload: {}'.format(payload)
    console_msg += '\n\nPosting to URI: {}'.format(uri)
    update_cell(cell=CONSOLE_CELL, value=console_msg, font_color=COLOR_BLUE)

def update_console_cmd_not_found(cmd_name):
    console_msg = 'Error, could not find: {}'.format(cmd_name),
    update_cell(cell=CONSOLE_CELL, value=console_msg, font_color=COLOR_BLUE)


# reset the status fields in the control panel
def reset_table_control_panel():
    for cell in CP_CELLS_TO_RESET:
        update_cell(cell=cell, value='')

# reset the status_code column for all tables
def reset_table_status_column():
    for table in TABLE_NAMES:
        print (table)
        t = xw.Range(table)
        for no, row in enumerate(t.value[2:]):
            cell = t[no + 2, 0].address
            xw.Range(cell).value = ''
            xw.Range(cell).color = COLOR_DEFAULT

def clear_status():
    reset_table_control_panel()
    reset_table_status_column()
