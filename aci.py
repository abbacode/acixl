import requests
import json
import sys
import collections
import excel
from requests.packages.urllib3.exceptions import InsecureRequestWarning

# Disable urllib3 warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


_missing = object()
class lazy_property(object):
    """
    Delays loading of property until first access. Credit goes to the
    Implementation in the werkzeug suite:
    http://werkzeug.pocoo.org/docs/utils/#werkzeug.utils.cached_property
    This should be used as a decorator in a class and in Evennia is
    mainly used to lazy-load handlers:
        ```python
        @lazy_property
        def attributes(self):
            return AttributeHandler(self)
        ```
    Once initialized, the `AttributeHandler` will be available as a
    property "attributes" on the object.
    """
    def __init__(self, func, name=None, doc=None):
        "Store all properties for now"
        self.__name__ = name or func.__name__
        self.__module__ = func.__module__
        self.__doc__ = doc or func.__doc__
        self.func = func

    def __get__(self, obj, type=None):
        "Triggers initialization"
        if obj is None:
            return self
        value = obj.__dict__.get(self.__name__, _missing)
        if value is _missing:
            value = self.func(obj)
        obj.__dict__[self.__name__] = value
        return value



# Class must be instantiated with APIC IP address, username, and password
# the login function returns the APIC c ookies.
class AciHandler(object):
    def __init__(self, apic='', user='', pword=''):
        self.apic = apic
        self.user = user
        self.pword = pword
        self.cookies = None

    @lazy_property
    def tenant_policies(self):
        return TenantHandler(self.apic,self.cookies)

    @lazy_property
    def fabric_policies(self):
        return FabricPolicyHandler(self.apic,self.cookies)

    @lazy_property
    def fabric_access(self):
        return FabricAccessHandler(self.apic,self.cookies)

    def login(self, show_auth_status=False):
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
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        cookies = None
        status = 0
        if show_auth_status:
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
        if show_auth_status:
            excel.update_auth_status(status)
        return status


# -----------------------------------
# All fabric related methods go here
# -----------------------------------
class FabricPolicyHandler(object):
    def __init__(self, apic, cookies):
        self.apic = apic
        self.cookies = cookies

class FabricAccessHandler(object):
    def __init__(self, apic, cookies):
        self.apic = apic
        self.cookies = cookies

# -----------------------------------
# All tenant related methods go here
# -----------------------------------
class TenantHandler(object):
    def __init__(self, apic, cookies):
        self.apic = apic
        self.cookies = cookies

    def create_tenant(self, **kwargs):
        payload = '''
        {{
            "fvTenant": {{
                "attributes": {{
                    "dn": "uni/tn-{tn_name}",
                    "name": "{tn_name}",
                    "descr":"{description}",
                    "status": "{action}"
                }}
            }}
        }}
        '''.format(**kwargs)

        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}.json'
                       .format(self.apic, kwargs['tn_name']),
                       data=json.dumps(payload),
                       cookies=self.cookies, verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status


    def create_vrf(self, **kwargs):
        payload = '''
        {{
            "fvCtx": {{
                "attributes": {{
                    "dn": "uni/tn-{tn_name}/ctx-{vrf_name}",
                    "knwMcastAct": "permit",
                    "name": "{vrf_name}",
                    "pcEnfPref": "{policy_enforce}",
                    "pcEnfDir": "{policy_direction}",
                    "descr":"{description}",
                    "status": "{action}"
                }}
            }}
        }}
        '''.format(**kwargs)
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}/ctx-{}.json'
                       .format(self.apic,
                               kwargs['tn_name'],
                               kwargs['vrf_name']),
                       data=json.dumps(payload),
                       cookies=self.cookies, verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Uknown Error'
        return status


    def create_anp(self, **kwargs):
        payload = '''
               {{
                   "fvAp": {{
                       "attributes": {{
                           "dn": "uni/tn-{tn_name}/ap-{anp_name}",
                           "name": "{anp_name}",
                           "status": "{action}"
                       }}
                   }}
               }}
        '''.format(**kwargs)
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}/ap-{}.json'
                       .format(self.apic,
                               kwargs['tn_name'],
                               kwargs['anp_name']),
                       data=json.dumps(payload),
                       cookies=self.cookies, verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status

    def create_bd(self, **kwargs):
        payload = '''
        {{
            "fvBD": {{
                "attributes": {{
                    "dn": "uni/tn-{tn_name}/BD-{bd_name}",
                    "mac": "00:22:BD:F8:19:FF",
                    "name": "{bd_name}",
                    "unicastRoute": "{unicast_routing}",
                    "arpFlood":"{arp_flooding}",
                    "unkMacUcastAct": "{l2_unknown_unicast}",
                    "unkMcastAct": "{l3_unknown_mcast}",
                    "multiDstPktAct": "{mdst_flooding}",
                    "limitIpLearnToSubnets":"{limit_iplearn_subnet}",
                    "descr":"{description}",
                    "status": "{action}"
                }},
                "children": [
                    {{
                        "fvRsCtx": {{
                            "attributes": {{
                                "tnFvCtxName": "{vrf_name}",
                                "status": "created,modified"
                            }}
                        }}
                    }}
                ]
            }}
        }}
        '''.format(**kwargs)
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}/BD-{}.json'
                       .format(self.apic,
                               kwargs['tn_name'],
                               kwargs['bd_name']),
                       data=json.dumps(payload),
                       cookies=self.cookies,
                       verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status

    def create_subnet(self, **kwargs):

        scope = ''
        if kwargs.get('private_to_vrf') == 'enabled':
            scope ='private,'
        if kwargs.get('advertised_externally') == 'enabled':
            scope ='public,'
        if kwargs.get('shared_between_vrfs') == 'enabled':
            scope+='shared'
        kwargs['scope'] = scope

        payload = '''
        {{
            "fvSubnet": {{
                "attributes": {{
                    "dn": "uni/tn-{tn_name}/BD-{bd_name}/subnet-[{subnet}]",
                    "ip": "{subnet}",
                    "scope": "{scope}",
                    "virtual": "{treat_as_virtual_ip}",
                    "preferred": "{make_primary_ip}",
                    "descr":"{description}",
                    "status": "{action}"
                }}
            }}
        }}
        '''.format(**kwargs)
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}/BD-{}/subnet-[{}].json'
                       .format(self.apic,
                               kwargs['tn_name'],
                               kwargs['bd_name'],
                               kwargs['subnet']),
                       data=json.dumps(payload),
                       cookies=self.cookies,
                       verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status


    def create_epg(self, **kwargs):
        payload = '''
        {{
            "fvAEPg": {{
                "attributes": {{
                    "dn": "uni/tn-{tn_name}/ap-{anp_name}/epg-{epg_name}",
                    "name": "{epg_name}",
                    "rn": "epg-{epg_name}",
                    "pcEnfPref": "{intra_epg_isolation}",
                    "status": "{action}"
                }},
                "children": [
                    {{
                        "fvRsBd": {{
                            "attributes": {{
                                "tnFvBDName": "{bd_name}",
                                "status": "created,modified"
                            }}
                        }}
                    }}
                ]
            }}
        }}
        '''.format(**kwargs)
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}/ap-{}/epg-{}.json'
                       .format(self.apic,
                               kwargs['tn_name'],
                               kwargs['anp_name'],
                               kwargs['epg_name']),
                       data=json.dumps(payload),
                       cookies=self.cookies,
                       verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status


    def create_epg_domain(self, **kwargs):
        payload = '''
        {{
            "fvRsDomAtt": {{
                "attributes": {{
                    "tDn": "uni/{domain_type}-{domain_name}",
                    "instrImedcy": "{deploy_immediacy}",
                    "resImedcy": "{resolution_immedacy}",
                    "status": "{action}"
                }}
            }}
        }}
        '''.format(**kwargs)
        payload = json.loads(payload, object_pairs_hook=collections.OrderedDict)
        print (payload)
        s = requests.Session()
        try:
            r = s.post('https://{}/api/node/mo/uni/tn-{}/ap-{}/epg-{}.json'
                       .format(self.apic,
                               kwargs['tn_name'],
                               kwargs['anp_name'],
                               kwargs['epg_name']),
                       data=json.dumps(payload),
                       cookies=self.cookies,
                       verify=False)
            status = r.status_code
        except Exception as e:
            status = 'Unknown Error'
        return status




# -----------------------------------
# ACI methods
# -----------------------------------
def create_all_tenants():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    if not aci.cookies:
        return        
    table = excel.get_table('TABLE_TENANT')
    for row in table:
        tenant = table[row]
        status_code = aci.tenant_policies.create_tenant(**tenant)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_TENANT',
                               last_action=' Update tenant configuration')


def create_all_vrfs():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    if not aci.cookies:
        return        
    table = excel.get_table('TABLE_VRF')
    for row in table:
        vrf = table[row]
        status_code = aci.tenant_policies.create_vrf(**vrf)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_VRF',
                               last_action=' Update VRF configuration')


def create_all_anps():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    table = excel.get_table('TABLE_ANP')
    if not aci.cookies:
        return        
    for row in table:
        anp = table[row]
        status_code = aci.tenant_policies.create_anp(**anp)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_ANP',
                               last_action=' Update application network profiles')


def create_all_bds():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    if not aci.cookies:
        return        
    table = excel.get_table('TABLE_BD')
    for row in table:
        bd = table[row]
        status_code = aci.tenant_policies.create_bd(**bd)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_BD',
                               last_action=' Update bridge domain configuration')

def create_all_subnets():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    if not aci.cookies:
        return        
    table = excel.get_table('TABLE_BD_SUBNET')
    for row in table:
        bd = table[row]
        status_code = aci.tenant_policies.create_subnet(**bd)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_BD_SUBNET',
                               last_action=' Update bridge domain subnet configuration')


def create_all_epgs():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    if not aci.cookies:
        return        
    table = excel.get_table('TABLE_EPG')
    for row in table:
        epg = table[row]
        status_code = aci.tenant_policies.create_epg(**epg)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_EPG',
                               last_action=' Update end point group configuration')



def create_all_epg_domains():
    aci = AciHandler(apic=excel.APIC, pword=excel.PWORD, user=excel.USER)
    aci.login(show_auth_status=True)
    if not aci.cookies:
        return        
    table = excel.get_table('TABLE_EPG_DOMAIN')
    for row in table:
        epg = table[row]
        status_code = aci.tenant_policies.create_epg_domain(**epg)
        excel.update_row_status(row, status_code)
    excel.update_script_status(table_name='TABLE_EPG_DOMAIN',
                               last_action=' Update end point group domain association')


# Initialize the fabric login method, passing appropriate variables

if __name__ == '__main__':
    print ('Do not execute this script manually')
    sys.exit()
