tmp_table_report
===
select idUserdata as [ntid] from user_data_mapping_role where idRegion='(%RG_NAME%)' and deleted=0
===
select bpRole.BpRoleStandardName as [Job Role] 
		from (user_data_mapping_role as udata   
			inner join BpRoleStandard as bpRole on udata.idBpRoleStandard = bpRole.id) 
		where udata.idUserdata = '(%VALUE%)' 
			and udata.idRegion='(%RG_NAME%)'
			and udata.Deleted = 0 and bpRole.Deleted = 0 
		group by bpRole.BpRoleStandardName
===
select UD.ntid, (UD.fname+','+UD.lname) As Fullname ,UD.fname,UD.lname,UD.omsSubfunction,tmp_table.[value],
	Cr.courseArena,rpc.courseId,
Cr.courseTitle,Cr.courseType,
Cr.courseDuration,rpc.ps,Cr.courseDelivery
	from (((user_data as UD 
		inner join 
	(select UDT.ntid, Course.courseId,CMR.ps
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
Course.deleted=0 and UMR.idRegion='(%RG_NAME%)' And UDT.Region='(%RG_NAME%)'
group by UDT.ntid,Course.courseId,CMR.ps
) as rpc on rpc.ntid = UD.ntid)
	inner join Course as Cr on Cr.courseId = rpc.courseId)
	inner join tmp_table_report as tmp_table on tmp_table.[key] = UD.ntid)
	order by rpc.ntid, rpc.courseId

