select rpc.ntid,rpc.Fullname,rpc.fname,rpc.lname,rpc.omsSubfunction from(
select UDT.ntid,(fname+','+lname) as Fullname,UDT.fname,
UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery,
(select count (*) from (select UDT.ntid,(fname+','+lname) as Fullname,
UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
from ((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
where UMR.deleted=0 and  UDT.deleted =0 and
BPROLE.deleted=0 and CMR.deleted=0 and
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
) as tbl_cached where tbl_cached.ntid = UDT.ntid and tbl_cached.courseId = Course.courseId )
 as count_conflict
from ((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
where UMR.deleted=0 and  UDT.deleted =0 and
BPROLE.deleted=0 and CMR.deleted=0 and
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
) as rpc
 where (rpc.count_conflict > 1 and rpc.ps = 'P') or rpc.count_conflict =1
=====
tmp_table_report
===
select idUserdata from  user_data_mapping_role where deleted = 0 and idFunction='(%RG_F_ID%)'
===
select bpRole.BpRoleStandardName as [Job Role] 
		from (user_data_mapping_role as udata   
			inner join BpRoleStandard as bpRole on udata.idBpRoleStandard = bpRole.id) 
		where udata.ntid = '(%VALUE%)' 
			and udata.idFunction='(%RG_F_ID%)'
			and udata.Deleted = 0 and bpRole.Deleted = 0 
		group by bpRole.BpRoleStandardName
===
select tmp_table.[value] AS [Job Role] from ((select rpc.ntid, rpc.courseArena,rpc.courseId,
rpc.courseTitle,rpc.courseType,rpc.courseDuration,rpc.ps,rpc.courseDelivery
from (
select UDT.ntid,(fname+','+lname) as Fullname,UDT.fname,
UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery,
(select count (*) from (select UDT.ntid,(fname+','+lname) as Fullname,
UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
from ((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
where UMR.deleted=0 and  UDT.deleted =0 and
BPROLE.deleted=0 and CMR.deleted=0 and
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
) as tbl_cached where tbl_cached.ntid = UDT.ntid and tbl_cached.courseId = Course.courseId )
 as count_conflict
from ((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
where UMR.deleted=0 and  UDT.deleted =0 and
BPROLE.deleted=0 and CMR.deleted=0 and
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
)
as rpc
where (rpc.count_conflict > 1 and rpc.ps = 'P') or rpc.count_conflict =1) as tbl_data inner join tmp_table_report as tmp_table on tmp_table.[key] = tbl_data.ntid)
=====
select rpc.courseArena,rpc.courseId,
rpc.courseTitle,rpc.courseType,
rpc.courseDuration,rpc.ps,rpc.courseDelivery from(
select UDT.ntid,(fname+','+lname) as Fullname,UDT.fname,
UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery,
(select count (*) from (select UDT.ntid,(fname+','+lname) as Fullname,
UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
from ((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
where UMR.deleted=0 and  UDT.deleted =0 and
BPROLE.deleted=0 and CMR.deleted=0 and
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
) as tbl_cached where tbl_cached.ntid = UDT.ntid and tbl_cached.courseId = Course.courseId )
 as count_conflict
from ((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
where UMR.deleted=0 and  UDT.deleted =0 and
BPROLE.deleted=0 and CMR.deleted=0 and
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
) as rpc
 where (rpc.count_conflict > 1 and rpc.ps = 'P') or rpc.count_conflict =1