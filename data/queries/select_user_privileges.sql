select p.ntid,r.RegionName,rmt.roleName,p.permission
from ((privileges as p
inner join Region as r
on p.idRegion = r.id)
inner join RMT_ROLES as rmt
on rmt.id = p.idRoleRMT)
where p.ntid = '(%VALUE%)' and p.deleted=0 and r.deleted=0 and rmt.deleted=0
group by p.ntid,r.RegionName,rmt.roleName,p.permission
order by p.ntid, r.RegionName, rmt.roleName