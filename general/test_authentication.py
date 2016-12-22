sys.path.insert(1, os.path.join(sys.path[0], '..'))
from aci import AciHandler
from excel import APIC, USER, PWORD

aci = AciHandler(APIC, USER, PWORD)
aci.login(show_auth_status=True)
