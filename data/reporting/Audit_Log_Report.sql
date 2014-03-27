select AD.ntid, AD.idFunction,AD.userAction, AD.data_fields, prev_value, new_value, table_name,AD.description,AD.timestamp
from audit_logs as AD
where AD.idFunction = '(%RG_NAME%)'
order by AD.timestamp DESC