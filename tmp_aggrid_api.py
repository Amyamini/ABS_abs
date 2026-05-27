import inspect
from st_aggrid import AgGrid
import st_aggrid
print('AgGrid signature:', inspect.signature(AgGrid))
print('st_aggrid attrs update:', [a for a in dir(st_aggrid) if 'Update' in a or 'ReturnMode' in a or 'GridUpdateMode' in a])
