sys.path.insert(1, os.path.join(sys.path[0], '..'))

from excel import set_table_action
from excel import TABLE_NAMES

for table in TABLE_NAMES:
    set_table_action(table, 'created')
