select p.ntid,r.RegionName,rmt.roleName,f.nameFunction,p.permission,f.id as [Function ID]
from ((((privileges as p
inner join Region_Function as rF
on p.idRegionFunction = rF.id)
inner join Functions as f
on rF.idFunctions = f.id)
inner join Region as r
on rF.idRegion = r.id)
inner join RMT_ROLES as rmt
on rmt.id = p.idRoleRMT)
where p.ntid = '(%VALUE%)' and p.deleted=0 and rF.deleted=0 and f.deleted=0 and r.deleted=0 and rmt.deleted=0
group by  p.ntid,r.RegionName,rmt.roleName,f.nameFunction,p.permission,f.id
order by p.ntid, r.RegionName, f.id, rmt.roleName