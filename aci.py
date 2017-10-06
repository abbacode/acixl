import requests
import json
import sys
import excel
import jinja2
from collections import OrderedDict
from requests.packages.urllib3.exceptions import InsecureRequestWarning

JSON_ROOT_FOLDER = 'C:\\acixl\\jsondata\\'
LAUNCH_COMMANDS = 'C:\\acixl\\launcher.json'
APIC_URI = 'https://{apic}/api/node/{payload_uri}.json'

# Disable urllib3 warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

class AciHandler(object):
    def __init__(self, apic='', user='', pword=''):
        self.apic = apic
        self.user = user
        self.pword = pword
        self.cookies = None
        self.commands = {}
        self.load_commands()

    def load_commands(self):
        try:
            with open(LAUNCH_COMMANDS, 'r') as f:
                self.commands = json.load(f)
        except Exception as e:
            self.commands = {}
            excel.update_console_cmd_not_found(LAUNCH_COMMANDS)

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
        excel.show_auth_attempt_msg()
        try:
            r = s.post('https://{}/api/mo/aaaLogin.json'.format(self.apic),
                       data=json.dumps(payload),
                       verify=False,
                       timeout=5)
            status = r.status_code
            self.cookies = r.cookies
        except Exception as e:
            status = 999
        excel.update_auth_status(status)
        return status

    def post(self, uri, payload):
        s = requests.Session()
        try:
            r = s.post(uri, data=payload, cookies=self.cookies, verify=False, timeout=5)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status


    def convert_bd_scope(self, row_data):
        scope = ''
        if row_data.get('private_to_vrf') == 'enabled':
            scope = 'private'
        if row_data.get('advertised_externally') == 'enabled':
            scope = 'public'
        if row_data.get('shared_between_vrfs') == 'enabled':
            scope += ',shared'
        return scope


    def push_to_apic(self, cmd):
        self.login()
        if not self.cookies:
            return
        if not self.commands:
            return
        # unpack command values from the dictionary
        script_msg = self.commands[cmd]['script_msg']
        table_name  =  self.commands[cmd]['table_name']
        json_folder = self.commands[cmd]['json_folder']
        json_file   = self.commands[cmd]['json_file']
        json_uri    = self.commands[cmd]['json_uri']

        # get data from the table in excel (i.e. TABLE_TENANT)
        table = excel.get_table(table_name)

        #load up the jinja2 crap
        template_loader = jinja2.FileSystemLoader(searchpath=(JSON_ROOT_FOLDER+json_folder))
        template_env = jinja2.Environment(loader=template_loader)
        template = template_env.get_template(json_file)

        for row in table:
            # convert scope for bd_subnet
            if 'bd_subnet' in cmd:
                scope = self.convert_bd_scope(table[row])
                table[row]['scope'] = scope

            # update the payload & uri values
            row_payload = template.render(**table[row])
            row_uri = str(json_uri.format(**table[row]))

            # generate the full URI to post the payload
            full_uri = APIC_URI.format(apic=self.apic, payload_uri=row_uri)

            # update the console cell in excel to show the output
            excel.update_console(row=int(row),table_name=table_name,
                                 uri=full_uri,payload=row_payload)

            # post the payload and get the status result
            status = self.post(full_uri,row_payload)

            #update the cell with the status result
            cell_to_update = table[row]['status_cell']
            excel.update_cell_status(cell_to_update, status)

        # update status for the overall exceution of the script
        excel.update_cp_status(table_name=table_name,script_msg=script_msg)

# This function is called from excel via xlwings addon
def run_from_excel(cmd):
    aci = AciHandler(excel.APIC,excel.USER,excel.PWORD)
    aci.push_to_apic(cmd)


if __name__ == '__main__':
    print ("This script should not be executed manually.")
    print ("Please perform all operations via the spreadsheet interface.")
    sys.exit()
