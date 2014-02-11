select rpc.ntid,rpc.Fullname,rpc.fname,rpc.lname,rpc.omsSubfunction 
 from (select udata.ntid, udata.region,(fname+','+lname) as Fullname,udata.fname,udata.lname,
udata.omsSubfunction,cou.courseArena,cou.courseId,cou.courseTitle,cou.courseType,
cou.courseDuration ,cou.courseDelivery,courseMapp.ps, 
(select count(*) from (select udata.ntid, udata.region,(fname+','+lname) as Fullname,udata.fname,udata.lname,
udata.omsSubfunction,cou.courseArena,cou.courseId,cou.courseTitle,cou.courseType,
cou.courseDuration ,cou.courseDelivery,courseMapp.ps
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse)
where cou.deleted=0 and courseMapp.deleted=0 and bpRole.deleted=0 and AcBpMapp.deleted=0
 and ac.deleted=0 and spAc.deleted=0 and sp.deleted=0 and udata.deleted=0 and
spAc.function_region='(%RG_F_ID%)' and udata.SFunction='(%RG_F_NAME%)'
group by udata.ntid,cou.courseId, cou.courseTitle, cou.courseType,
udata.fname,cou.courseDuration,cou.courseDelivery,cou.courseArena,cou.id,
udata.lname,udata.omsSubfunction,courseMapp.ps,udata.region) as tbl_cached where tbl_cached.ntid = udata.ntid and tbl_cached.courseId=cou.courseId
) as count_conflict
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse)
where cou.deleted=0 and courseMapp.deleted=0 and bpRole.deleted=0 and AcBpMapp.deleted=0
 and ac.deleted=0 and spAc.deleted=0 and sp.deleted=0 and udata.deleted=0 and
spAc.function_region='(%RG_F_ID%)' and udata.SFunction='(%RG_F_NAME%)'
group by udata.ntid,cou.courseId, cou.courseTitle, cou.courseType,
udata.fname,cou.courseDuration,cou.courseDelivery,cou.courseArena,cou.id,
udata.lname,udata.omsSubfunction,courseMapp.ps,udata.region) as rpc
 where (rpc.count_conflict > 1 and rpc.ps = 'P') or rpc.count_conflict =1
=====
tmp_table_report
===
select ntid from user_data where deleted = 0 and SFunction='(%RG_F_NAME%)'
===
select bpRole.BpRoleStandardName as [Job Role] 
		from (((((((user_data as udata  
			inner join specialism as sp on sp.SpecialismName = udata.specialism) 
			inner join SpecialismMappingActivity as spAc on sp.id = spAc.idSpecialism) 
			inner join Activity as ac on spAc.idActivity =ac.id) 
			inner join MappingActivityBpStandardRole as AcBpMapp on ac.id = AcBpMapp.idActivity) 
			inner join BpRoleStandard as bpRole on AcBpMapp.idBpRoleStandard = bpRole.id) 
			inner join CourseMappingBpRoleStandard as courseMapp on courseMapp.idBpRole = bpRole.id) 
			inner join course as cou on cou.id = courseMapp.idCourse ) 
		where udata.ntid='(%VALUE%)' 
			and udata.SFunction='(%RG_F_NAME%)'
			and udata.Deleted = 0 and spAc.Deleted=0 
			and sp.Deleted=0 and ac.Deleted = 0 
			and AcBpMapp.Deleted = 0 
			and bpRole.Deleted = 0 
			and courseMapp.Deleted =0
			and spAc.function_region='(%RG_F_ID%)'
		group by bpRole.BpRoleStandardName
===
select tmp_table.[value] AS [Job Role] from ((select rpc.ntid, rpc.courseArena,rpc.courseId,
rpc.courseTitle,rpc.courseType,rpc.courseDuration,rpc.ps,rpc.courseDelivery
 from (select udata.ntid, udata.region,(fname+','+lname) as Fullname,udata.fname,udata.lname,
udata.omsSubfunction,cou.courseArena,cou.courseId,cou.courseTitle,cou.courseType,
cou.courseDuration ,cou.courseDelivery,courseMapp.ps, 
(select count(*) from (select udata.ntid, udata.region,(fname+','+lname) as Fullname,udata.fname,udata.lname,
udata.omsSubfunction,cou.courseArena,cou.courseId,cou.courseTitle,cou.courseType,
cou.courseDuration ,cou.courseDelivery,courseMapp.ps
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse)
where cou.deleted=0 and courseMapp.deleted=0 and bpRole.deleted=0 and AcBpMapp.deleted=0
 and ac.deleted=0 and spAc.deleted=0 and sp.deleted=0 and udata.deleted=0 and
spAc.function_region='(%RG_F_ID%)' and udata.SFunction='(%RG_F_NAME%)'
group by udata.ntid,cou.courseId, cou.courseTitle, cou.courseType,
udata.fname,cou.courseDuration,cou.courseDelivery,cou.courseArena,cou.id,
udata.lname,udata.omsSubfunction,courseMapp.ps,udata.region) as tbl_cached where tbl_cached.ntid = udata.ntid and tbl_cached.courseId=cou.courseId
) as count_conflict
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse)
where cou.deleted=0 and courseMapp.deleted=0 and bpRole.deleted=0 and AcBpMapp.deleted=0
 and ac.deleted=0 and spAc.deleted=0 and sp.deleted=0 and udata.deleted=0 and
spAc.function_region='(%RG_F_ID%)' and udata.SFunction='(%RG_F_NAME%)'
group by udata.ntid,cou.courseId, cou.courseTitle, cou.courseType,
udata.fname,cou.courseDuration,cou.courseDelivery,cou.courseArena,cou.id,
udata.lname,udata.omsSubfunction,courseMapp.ps,udata.region) as rpc
 where (rpc.count_conflict > 1 and rpc.ps = 'P') or rpc.count_conflict =1) as tbl_data inner join tmp_table_report as tmp_table on tmp_table.[key] = tbl_data.ntid)
=====
select rpc.courseArena,rpc.courseId,
rpc.courseTitle,rpc.courseType,rpc.courseDuration,rpc.ps,rpc.courseDelivery
 from (select udata.ntid, udata.region,(fname+','+lname) as Fullname,udata.fname,udata.lname,
udata.omsSubfunction,cou.courseArena,cou.courseId,cou.courseTitle,cou.courseType,
cou.courseDuration ,cou.courseDelivery,courseMapp.ps, 
(select count(*) from (select udata.ntid, udata.region,(fname+','+lname) as Fullname,udata.fname,udata.lname,
udata.omsSubfunction,cou.courseArena,cou.courseId,cou.courseTitle,cou.courseType,
cou.courseDuration ,cou.courseDelivery,courseMapp.ps
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse)
where cou.deleted=0 and courseMapp.deleted=0 and bpRole.deleted=0 and AcBpMapp.deleted=0
 and ac.deleted=0 and spAc.deleted=0 and sp.deleted=0 and udata.deleted=0 and
spAc.function_region='(%RG_F_ID%)' and udata.SFunction='(%RG_F_NAME%)'
group by udata.ntid,cou.courseId, cou.courseTitle, cou.courseType,
udata.fname,cou.courseDuration,cou.courseDelivery,cou.courseArena,cou.id,
udata.lname,udata.omsSubfunction,courseMapp.ps,udata.region) as tbl_cached where tbl_cached.ntid = udata.ntid and tbl_cached.courseId=cou.courseId
) as count_conflict
from (((((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.id)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.id = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
inner join CourseMappingBpRoleStandard as courseMapp
on courseMapp.idBpRole = bpRole.id)
inner join course as cou
on cou.id = courseMapp.idCourse)
where cou.deleted=0 and courseMapp.deleted=0 and bpRole.deleted=0 and AcBpMapp.deleted=0
 and ac.deleted=0 and spAc.deleted=0 and sp.deleted=0 and udata.deleted=0 and
spAc.function_region='(%RG_F_ID%)' and udata.SFunction='(%RG_F_NAME%)'
group by udata.ntid,cou.courseId, cou.courseTitle, cou.courseType,
udata.fname,cou.courseDuration,cou.courseDelivery,cou.courseArena,cou.id,
udata.lname,udata.omsSubfunction,courseMapp.ps,udata.region) as rpc
 where (rpc.count_conflict > 1 and rpc.ps = 'P') or rpc.count_conflict =1