# APIC authentication cell locations
APIC = Cell('APIC','B5').value
USER = Cell('APIC','B6').value
PWORD = Cell('APIC','B7').value

# Control panel cell locations
AUTH_STATUS_1 = 'B3'
AUTH_STATUS_2 = 'B4'
SCRIPT_STATUS_1 = 'C5'
SCRIPT_STATUS_2 = 'C6'

# Control panel cells to reset when button is pushed
CP_CELLS_TO_RESET = (AUTH_STATUS_1, SCRIPT_STATUS_1,
                     AUTH_STATUS_2, SCRIPT_STATUS_2)

# Table named ranges used in excel
TABLE_NAMES = ('TABLE_TENANT','TABLE_VRF','TABLE_ANP',
               'TABLE_BD','TABLE_EPG','TABLE_EPG_DOMAIN',
               'TABLE_BD_SUBNET')

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
                  'shared_between_vrfs': 'false',
                  'intra_epg_isolation': 'unenforced'}

# Cells that are mandatory, otherwise ignore the row
MANDATORY_VALUES = ('tn_name', 'anp_name', 'bd_name',
                    'vrf_name', 'subnet')

# Cell color codes
COLOR_DEFAULT = 'FDFEFE'
COLOR_FAILED = 'E74C3C'
COLOR_PASS = '58D68D'
COLOR_IGNORED ='F0B27A'
COLOR_BLACK = 'black'


# update the value and color of a cell
def update_cell(cell='', text='', color=COLOR_DEFAULT, font_color='black'):
    Cell(cell).value = text
    Cell(cell).color = color
    Cell(cell).font.color = font_color


# update the status_code column for a row
def update_row_status(row, status_code):
    cell = Cell(row,1)
    cell.value = status_code
    cell.color = get_status_color (status_code)

# update the script status field whenever a script is executed
def update_script_status(table_name, last_action='N/A'):
    table_status = get_table_status(table_name)
    # update the cell to reference what action was executed
    update_cell(cell=SCRIPT_STATUS_1,
                text=last_action,
                color=COLOR_DEFAULT)
    # update the cell to indicate how successful it was
    if all(status == 200 for status in table_status):
        update_cell(cell=SCRIPT_STATUS_2,
                    text=' Complete success',
                    color=COLOR_DEFAULT,
                    font_color='green')
    elif any(status == 200 for status in table_status):
        update_cell(cell=SCRIPT_STATUS_2,
                    text=' Partial success',
                    color=COLOR_DEFAULT,
                    font_color='E67E22')
    else:
        update_cell(cell=SCRIPT_STATUS_2,
                    text=' Complete failure',
                    color=COLOR_DEFAULT,
                    font_color='red')


def get_status_color(status_code):
    COLORS = {200: COLOR_PASS,
              'Aborted - missing fields': COLOR_IGNORED}
    color = COLORS.get(status_code, COLOR_FAILED)
    return color


def show_auth_attempt_msg():
    update_cell(cell=AUTH_STATUS_1,
                text='Attempting authentication...')
    update_cell(cell=AUTH_STATUS_2,
                text='Do not push policies unless this works')

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
                    'msg2': 'Check network connectivity',
                    'color': COLOR_FAILED}}

    if STATUS.get(status_code):
        msg_1 = STATUS.get(status_code)['msg1']
        msg_2 = STATUS.get(status_code)['msg2']
        status_color = STATUS.get(status_code)['color']
        update_cell(cell=AUTH_STATUS_1, text=msg_1, color=status_color)
        update_cell(cell=AUTH_STATUS_2, text=msg_2, color=status_color)
        if status_code != 200:
            update_cell(cell=SCRIPT_STATUS_1, text='Script Aborted',
                        font_color=COLOR_FAILED)
            update_cell(cell=SCRIPT_STATUS_2, text=' ',
                        font_color=COLOR_FAILED)



# reset the status fields in the control panel
def reset_table_control_panel():
    for cell in CP_CELLS_TO_RESET:
        update_cell(cell=cell, text=' ')


# reset the status_code column for all tables
def reset_table_status_column():
    for table in TABLE_NAMES:
        table_rows = get_table_rows(table, skip_first_row=True)
        for row in table_rows:
            update_cell(cell=(row,1), text=' ')


# change the action field for a table
def set_table_action(table_name, action):
    table_rows = get_table_rows(table_name, skip_first_row=True)
    for row in table_rows:
        update_cell(cell=(row,2), text=action)


# return a list of column numbers used by a table
def get_table_cols(table_name, skip_first_col=False):
    table_cols = [cell.position[1] for cell in CellRange(table_name)]
    table_cols = list(sorted(set(table_cols)))
    if skip_first_col:
        table_cols.pop(0)
    return table_cols


# return a list of row numbers used by a table
def get_table_rows(table_name, skip_first_row=False):
    table_rows = [cell.position[0] for cell in CellRange(table_name)]
    table_rows = list(sorted(set(table_rows)))
    if skip_first_row:
        table_rows.pop(0)
    return table_rows


# return a list of headings used by a table
def get_table_headings(table_name, skip_first_heading=False):
    table_rows = get_table_rows(table_name)
    table_columns = get_table_cols(table_name)
    heading_row = table_rows[0]-1
    #print ('table_rows: {}'.format(table_rows))
    #print ('heading_row: {}'.format(heading_row))
    headings = [Cell(heading_row, col).value for col in table_columns]
    if skip_first_heading:
        headings.pop(0)
    return headings

# return a list of status codes for all rows in a table
def get_table_status(table_name):
    table_rows = get_table_rows(table_name, skip_first_row=True)
    status_codes = [Cell(row,1).value for row in table_rows]
    return status_codes


# remove rows that do not have mandatory values from a table
def remove_invalid_rows(table):
    temp_table = dict(table)
    for row in sorted(temp_table):
        for key in temp_table[row].keys():
            if key in MANDATORY_VALUES:
                value = temp_table[row][key]
                if not value:
                    if table.get(row):
                        #print('mandatory cell missing, deleting row: {}'.format(row))
                        del table[row]
                    update_cell(cell=(row,1),
                                text='Aborted - missing fields',
                                color=COLOR_IGNORED)
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


# return a dictionary using row number and column heading as the key/sub keys for a table
def get_table(table_name):
    table_rows = get_table_rows(table_name)
    table_columns = get_table_cols(table_name)
    table_headings = get_table_headings(table_name)
    table = {k: {} for k in table_rows}

    for row in table_rows:
        for col in table_columns:
            heading = table_headings[col-1]
            value = Cell(row,col).value
            #print ('row: {}, heading_pos: {}, heading: {} and value: {}'.format(row,col,heading,value))
            table[row][heading] = value

    # pop the first row, it's informational only
    table.pop(table_rows[0])

    # remove invalid rows
    table = remove_invalid_rows(table)

    # if a non-mandatory cell has no value, then use default values
    table = apply_default_values(table)

    return table


def set_table_default_values():
    for table in TABLE_NAMES:
        table_rows = get_table_rows(table, skip_first_row=True)
        table_cols = get_table_cols(table, skip_first_col=True)
        headings = get_table_headings(table, skip_first_heading=True)
        for row in table_rows:
            for col in table_cols:
                cell = Cell(row,col)
                value = cell.value
                if not value:
                    h = headings[col-2]
                    default_value = DEFAULT_VALUES.get(h)
                    update_cell(cell=cell.name, text=default_value)
                    #print (' -- for row {} -- heading searched: {}'.format(row,h))
                    #print ('   -- current value: {}'.format(value))
                    #print ('   -- empty value changed to : {}'.format(DEFAULT_VALUES.get(h)))