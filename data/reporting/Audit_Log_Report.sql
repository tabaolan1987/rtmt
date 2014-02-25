select AD.ntid,F.nameFunction,AD.userAction,AD.description,AD.timestamp
from audit_logs as AD
inner join functions as F
on AD.idFunction = F.id
where AD.idFunction = '(%RG_F_ID%)'