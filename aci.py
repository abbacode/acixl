import requests
import json
import sys
import excel
import jinja2
from collections import OrderedDict
from requests.packages.urllib3.exceptions import InsecureRequestWarning

commands = {'tenant': {'json_folder': 'FabTnPol',
                       'json_file'  : 'tenant.json',
                       'json_uri': 'mo/uni/tn-{tn_name}',
                       'table_name': 'TABLE_TENANT',
                       'script_msg': 'Push tenant configuration',
                       },
            'vrf': {'json_folder': 'FabTnPol',
                    'json_file': 'vrf.json',
                    'json_uri': 'mo/uni/tn-{tn_name}/ctx-{vrf_name}',
                    'table_name': 'TABLE_VRF',
                    'script_msg': 'Push VRF configuration',
                    },
            'app_profile': {'json_folder': 'FabTnPol',
                            'json_file': 'app_profile.json',
                            'json_uri': 'mo/uni/tn-{tn_name}/ap-{anp_name}',
                            'table_name': 'TABLE_ANP',
                            'script_msg': 'Push app profile configuration',
                            },
            'bd': {'json_folder': 'FabTnPol',
                   'json_file': 'bd.json',
                   'json_uri': 'mo/uni/tn-{tn_name}/BD-{bd_name}',
                   'table_name': 'TABLE_BD',
                   'script_msg': 'Push bridge domain configuration',
                   },
            'bd_subnet': {'json_folder': 'FabTnPol',
                          'json_file': 'bd_subnet.json',
                          'json_uri': 'mo/uni/tn-{tn_name}/BD-{bd_name}',
                          'table_name': 'TABLE_BD_SUBNET',
                          'script_msg': 'Push BD subnet configuration',
                          },
            'epg': {'json_folder': 'FabTnPol',
                    'json_file': 'epg.json',
                    'json_uri': 'mo/uni/tn-{tn_name}/ap-{anp_name}/epg-{epg_name}',
                    'table_name': 'TABLE_EPG',
                    'script_msg': 'Push EPG configuration',
                    },

            }

# Disable urllib3 warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# Top folder containing all the JSON files
JSON_ROOT = 'C:\\acixl\\jsondata\\'

#URI format before values are unpacked
APIC_URI = 'https://{apic}/api/node/{payload_uri}.json'

class AciHandler(object):
    def __init__(self, apic='', user='', pword=''):
        self.apic = apic
        self.user = user
        self.pword = pword
        self.cookies = None

    def login(self, show_auth_attempt=True):
        # Load login json payload
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
        if show_auth_attempt:
            excel.show_auth_attempt_msg()

        # Try the request, if exception, exit program w/ error
        try:
            # Verify is disabled as there are issues if it is enabled
            r = s.post('https://{}/api/mo/aaaLogin.json'.format(self.apic),
                       data=json.dumps(payload), verify=False)
            # Capture HTTP status code from the request
            status = r.status_code
            # Capture the APIC cookie for all other future calls
            self.cookies = r.cookies
        except Exception as e:
            status = 999
        if show_auth_attempt:
            excel.update_auth_status(status)
        return status

    def post(self, uri, payload):
        s = requests.Session()
        try:
            r = s.post(uri, data=payload, cookies=self.cookies,verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status


    def push_to_apic(self, cmd):
        self.login()
        if not self.cookies:
            return

        # unpack command values from the dictionary
        script_msg = commands[cmd]['script_msg']
        table_name  =  commands[cmd]['table_name']
        json_folder = commands[cmd]['json_folder']
        json_file   = commands[cmd]['json_file']
        json_uri    = commands[cmd]['json_uri']

        # get data from the table in excel (i.e. TABLE_TENANT)
        table = excel.get_table(table_name)

        #load up the jinja2 crap
        template_loader = jinja2.FileSystemLoader(searchpath=(JSON_ROOT+json_folder))
        template_env = jinja2.Environment(loader=template_loader)
        template = template_env.get_template(json_file)


        # loop over all the rows in the table
        for row in table:

            # convert scope for bd_subnet
            if 'bd_subnet' in cmd:
                scope = ''
                if table[row].get('private_to_vrf') == 'enabled':
                    scope = 'private'
                if table[row].get('advertised_externally') == 'enabled':
                    scope = 'public'
                if table[row].get('shared_between_vrfs') == 'enabled':
                    scope += ',shared'
                table[row]['scope'] = scope

            # update the payload & uri values
            row_payload = template.render(**table[row])
            row_uri = str(json_uri.format(**table[row]))
            # generate the full URI to post the payload
            full_uri = APIC_URI.format(apic=self.apic, payload_uri=row_uri)

            #if DEBUG:
            #    print (row_payload)
            #    print (row_uri)
            #    print (full_uri)

            # post the payload and get the status result
            status = self.post(full_uri,row_payload)

            #if DEBUG:
            #    print (status)

            #update the cell with the status result
            cell_to_update = table[row]['status_cell']
            excel.update_cell_status(cell_to_update, status)

        excel.update_script_status(table_name=table_name,script_msg=script_msg)
        #if DEBUG:
        #    wait()


# This function is called from excel via xlwings addon
def run_from_excel(cmd):
    aci = AciHandler(excel.APIC,excel.USER,excel.PWORD)
    aci.push_to_apic(cmd)


# Initialize the fabric login method, passing appropriate variables

if __name__ == '__main__':
    print ('There is no need to execute this script manually.')
    print ('Open the excel sheet and navigate using the avaialble buttons.')
    sys.exit()
