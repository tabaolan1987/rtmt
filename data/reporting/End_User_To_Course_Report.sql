select rpc.ntid,rpc.Fullname,rpc.fname,rpc.lname,rpc.omsSubfunction from(
select UDT.ntid,(fname+','+lname) as Fullname,UDT.fname,
UDT.lname,UDT.omsSubfunction,
Course.courseId,
CMR.ps,
(select count (*) from (select UDT1.ntid,(fname+','+lname) as Fullname,
UDT1.fname,UDT1.lname,UDT1.omsSubfunction,
Course1.courseId,
CMR1.ps
from ((((user_data_mapping_role as UMR1
inner join user_data as UDT1
on UMR1.idUserdata = UDT1.ntid)
inner join BpRoleStandard as BPROLE1
on UMR1.idBpRoleStandard = BPROLE1.id)
inner join CourseMappingBpRoleStandard as CMR1
on UMR1.idBpRoleStandard = CMR1.idBpRole)
inner join Course as Course1
on CMR1.idCourse = Course1.id)
where UMR1.deleted=0 and  UDT1.deleted =0 and
BPROLE1.deleted=0 and CMR1.deleted=0 and
Course1.deleted=0 and UMR1.idFunction='(%RG_F_ID%)' And UDT1.SFunction='(%RG_F_NAME%)'
group by UDT1.ntid,UDT1.fname,UDT1.lname,UDT1.omsSubfunction,
Course1.courseId,
CMR1.ps
order by UDT1.ntid,Course1.courseId
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
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)' and UDT.SFunction='(%RG_F_NAME%)'
group by UDT.ntid,UDT.fname,UDT.lname,UDT.omsSubfunction,
Course.courseId,
CMR.ps
order by UDT.ntid,Course.courseId
) as rpc

=====
tmp_table_report
===
select ntid from user_data where deleted = 0 and SFunction='(%RG_F_NAME%)' 
===
select bpRole.BpRoleStandardName as [Job Role] 
		from (user_data_mapping_role as udata   
			inner join BpRoleStandard as bpRole on udata.idBpRoleStandard = bpRole.id) 
		where udata.idUserdata = '(%VALUE%)' 
			and udata.idFunction='(%RG_F_ID%)'
			and udata.Deleted = 0 and bpRole.Deleted = 0 
		group by bpRole.BpRoleStandardName
===
select tmp_table.[value] AS [Job Role] from ((select rpc.ntid, rpc.courseArena,rpc.courseId,
rpc.courseTitle,rpc.courseType,rpc.courseDuration,rpc.ps,rpc.courseDelivery
from (
select UDT.ntid,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery,
(select count (*) from (select UDT1.ntid,
Course1.courseArena,Course1.courseId,
Course1.courseTitle,Course1.courseType,
Course1.courseDuration,CMR1.ps,Course1.courseDelivery
from ((((user_data_mapping_role as UMR1
inner join user_data as UDT1
on UMR1.idUserdata = UDT1.ntid)
inner join BpRoleStandard as BPROLE1
on UMR1.idBpRoleStandard = BPROLE1.id)
inner join CourseMappingBpRoleStandard as CMR1
on UMR1.idBpRoleStandard = CMR1.idBpRole)
inner join Course as Course1
on CMR1.idCourse = Course1.id)
where UMR1.deleted=0 and  UDT1.deleted =0 and
BPROLE1.deleted=0 and CMR1.deleted=0 and
Course1.deleted=0 and UMR1.idFunction='(%RG_F_ID%)' and UDT1.SFunction='(%RG_F_NAME%)'
group by UDT1.ntid,
Course1.courseArena,Course1.courseId,
Course1.courseTitle,Course1.courseType,
Course1.courseDuration,CMR1.ps,Course1.courseDelivery
order by UDT1.ntid,Course1.courseId
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
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)' and UDT.SFunction='(%RG_F_NAME%)'
group by UDT.ntid,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
order by UDT.ntid,Course.courseId
)
as rpc
) as tbl_data inner join tmp_table_report as tmp_table on tmp_table.[key] = tbl_data.ntid)
order by tbl_data.ntid,tbl_data.courseId
=====
select rpc.courseArena,rpc.courseId,
rpc.courseTitle,rpc.courseType,
rpc.courseDuration,rpc.ps,rpc.courseDelivery,rpc.count_conflict from(
select UDT.ntid,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery,
(select count (*) from (select UDT1.ntid,
Course1.courseArena,Course1.courseId,
Course1.courseTitle,Course1.courseType,
Course1.courseDuration,CMR1.ps,Course1.courseDelivery
from ((((user_data_mapping_role as UMR1
inner join user_data as UDT1
on UMR1.idUserdata = UDT1.ntid)
inner join BpRoleStandard as BPROLE1
on UMR1.idBpRoleStandard = BPROLE1.id)
inner join CourseMappingBpRoleStandard as CMR1
on UMR1.idBpRoleStandard = CMR1.idBpRole)
inner join Course as Course1
on CMR1.idCourse = Course1.id)
where UMR1.deleted=0 and  UDT1.deleted =0 and
BPROLE1.deleted=0 and CMR1.deleted=0 and
Course1.deleted=0 and UMR1.idFunction='(%RG_F_ID%)' And UDT1.SFunction='(%RG_F_NAME%)'
group by UDT1.ntid,
Course1.courseArena,Course1.courseId,
Course1.courseTitle,Course1.courseType,
Course1.courseDuration,CMR1.ps,Course1.courseDelivery
order by UDT1.ntid,Course1.courseId
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
Course.deleted=0 and UMR.idFunction='(%RG_F_ID%)' And UDT.SFunction='(%RG_F_NAME%)'
group by UDT.ntid,
Course.courseArena,Course.courseId,
Course.courseTitle,Course.courseType,
Course.courseDuration,CMR.ps,Course.courseDelivery
order by UDT.ntid,Course.courseId
) as rpc
