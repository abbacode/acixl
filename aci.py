import requests
import json
import excel
import jinja2
from collections import OrderedDict
from requests.packages.urllib3.exceptions import InsecureRequestWarning

JSON_ROOT_FOLDER = 'C:\\acixl\\jsondata\\'
LAUNCHER_FILE = 'C:\\acixl\\launcher.json'
APIC_URI = 'https://{apic}/api/node/{payload_uri}.json'

# Disable urllib3 warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

class LaunchFileHandler(object):
    def __init__(self):
        self.data = self.read_data_from_file()

    def read_data_from_file(self):
        try:
            with open(LAUNCHER_FILE, 'r') as f:
                return json.load(f)
        except Exception as e:
            excel.show_console_launcher_error(launcher_fname=LAUNCHER_FILE)
            return {}

    @property
    def command_list(self):
        return list(self.data.keys()) if self.data else {}

    @property
    def table_list(self):
        worksheets = []
        tables = []
        for cmd in self.data:
            worksheets.append(self.data[cmd].get('worksheet_name'))
            tables.append(self.data[cmd].get('table_name'))
        combined_list = [worksheets]+[tables]
        return combined_list


class AciHandler(object):
    def __init__(self, apic='', user='', pword=''):
        self.apic = apic
        self.user = user
        self.pword = pword
        self.cookies = None
        self.launcher = LaunchFileHandler()

    def login(self):
        payload = '''
        {{
            "aaaUser": {{
                "attributes": {{
                    "name": "{user}",
                    "pwd": "{pword}"
                }}
            }}
        }}
        '''.format(user=self.user, pword=self.pword)
        payload = json.loads(payload, object_pairs_hook=OrderedDict)
        s = requests.Session()
        excel.show_cp_authentication_attempt_msg()
        try:
            uri = 'https://{}/api/mo/aaaLogin.json'.format(self.apic)
            r = s.post(uri,data=json.dumps(payload), verify=False, timeout=5)
            status = r.status_code
            self.cookies = r.cookies
        except Exception as e:
            status = 999
        excel.update_cp_authentication_response(status)
        return status

    def post(self, uri, payload):
        s = requests.Session()
        try:
            r = s.post(uri, data=payload, cookies=self.cookies, verify=False, timeout=5)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status


    def format_bd_scope(self, row_data):
        scope = ''
        if row_data.get('private_to_vrf') == 'enabled':
            scope = 'private'
        if row_data.get('advertised_externally') == 'enabled':
            scope = 'public'
        if row_data.get('shared_between_vrfs') == 'enabled':
            scope += ',shared'
        return scope


    def push_to_apic(self, cmd):
        # unpack command values from the dictionary
        json_folder = self.launcher.data[cmd]['json_folder']
        json_file = self.launcher.data[cmd]['json_file']
        json_uri = self.launcher.data[cmd]['json_uri']
        action_msg = self.launcher.data[cmd]['action_msg']
        table_name = self.launcher.data[cmd]['table_name']
        mandatory_keys = self.launcher.data[cmd]['mandatory_keys']
        default_values = self.launcher.data[cmd]['default_values']

        # get data from the table in excel (i.e. TABLE_TENANT)
        table = excel.get_table(table_name=table_name,
                                mandatory_keys=mandatory_keys,
                                default_values=default_values)

        #load up the jinja2 crap
        template_loader = jinja2.FileSystemLoader(searchpath=(JSON_ROOT_FOLDER+
                                                              json_folder))
        template_env = jinja2.Environment(loader=template_loader)
        template = template_env.get_template(json_file)

        for row in table:
            # convert scope for bd_subnet
            if 'bd_subnet' in cmd:
                scope = self.format_bd_scope(table[row])
                table[row]['scope'] = scope

            # update the payload & uri values
            row_payload = template.render(**table[row])
            row_uri = str(json_uri.format(**table[row]))

            # generate the full URI to post the payload
            full_uri = APIC_URI.format(apic=self.apic, payload_uri=row_uri)

            # update the console cell in excel to show the output
            excel.show_console_payload(row=int(row),table_name=table_name,
                                         uri=full_uri,payload=row_payload)

            # post the payload and get the status result
            status = self.post(full_uri,row_payload)

            #update the cell with the status result
            row_status_location = table[row]['status_cell']

            excel.update_status(row_status_location, status)

        # update status for the overall exceution of the script
        excel.show_push_report_status(table_name=table_name,
                                      action_msg=action_msg)


# This function is called from excel via xlwings addon
def run_from_excel(cmd):
    aci = AciHandler(apic=excel.APIC,user=excel.USER,pword=excel.PWORD)
    aci.login()
    if not aci.cookies:
        return
    if not aci.launcher.data:
        return
    if not aci.launcher.data.get(cmd):
        return
    if not excel.can_run_cmd_from_worksheet(
            cmd=cmd,launcher_data=aci.launcher.data):
        return
    aci.push_to_apic(cmd)

def refresh_excel_data():
    """
    Used to update the hidden _commands and _tables worksheet
    in excel with the data extracted from the launcher.json file.

    This function is executed when the RefreshCmd

    """
    json_file = LaunchFileHandler()
    if not json_file.data:
        return
    excel.update_excel_data(commands=json_file.command_list,
                            tables=json_file.table_list)


if __name__ == '__main__':
    print ("This script should not be executed manually.")
    print ("Please perform all operations via the spreadsheet interface.")

