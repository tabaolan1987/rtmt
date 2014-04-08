tmp_table_report
===
select distinct UM.idUserdata as [value] from (user_data_mapping_role AS UM inner join user_data AS U
					on U.ntid = UM.idUserdata)
	where UM.idRegion='(%RG_NAME%)' and UM.deleted=0 and U.deleted=0 and U.Region='(%RG_NAME%)' and U.suspend=0
		and U.SFunction (%CUSTOM_FILTER_NAME%)
===
select bpRole.BpRoleStandardName as [Job Role] 
		from ((user_data_mapping_role as udata   
			inner join BpRoleStandard as bpRole on udata.idBpRoleStandard = bpRole.id) 
			inner join user_data AS UD on UD.ntid = udata.idUserdata)
		where udata.idUserdata = '(%VALUE%)' 
			and udata.idRegion='(%RG_NAME%)'
			and udata.Deleted = 0 and bpRole.Deleted = 0 and UD.deleted=0
			and UD.Region='(%RG_NAME%)'
			and UD.suspend=0
			and UD.SFunction (%CUSTOM_FILTER_NAME%)
		group by bpRole.BpRoleStandardName
===
select UD.ntid, (UD.fname+','+UD.lname) As Fullname ,UD.fname,UD.lname,UD.omsSubfunction,tmp_table.[value],
	Cr.courseArena,rpc.courseId,
Cr.courseTitle,Cr.courseType,
Cr.courseDuration,rpc.ps,Cr.courseDelivery
	from (((user_data as UD 
		inner join 
	(select UDT.ntid, Course.courseId,CMR.ps, F.id As Fid
from (((((user_data_mapping_role as UMR
inner join user_data as UDT
on UMR.idUserdata = UDT.ntid)
inner join BpRoleStandard as BPROLE
on UMR.idBpRoleStandard = BPROLE.id)
inner join CourseMappingBpRoleStandard as CMR
on UMR.idBpRoleStandard = CMR.idBpRole)
inner join Course as Course
on CMR.idCourse = Course.id)
inner join Functions As F
on F.nameFunction = UDT.SFunction)
where UMR.deleted=0 and  UDT.deleted =0 and UDT.suspend=0
and BPROLE.deleted=0 and CMR.deleted=0
and F.deleted=0
and Course.deleted=0 and UMR.idRegion='(%RG_NAME%)' And UDT.Region='(%RG_NAME%)'
and UDT.SFunction (%CUSTOM_FILTER_NAME%)
and CMR.idRegion='(%RG_NAME%)'
and CMR.idFunction = F.id
and CMR.idFunction (%CUSTOM_FILTER_ID%)
and Course.idRegion='(%RG_NAME%)'
and Course.idFunction (%CUSTOM_FILTER_ID%)
and Course.idFunction = F.id
group by UDT.ntid,Course.courseId,CMR.ps, F.id
) as rpc on rpc.ntid = UD.ntid)
	inner join Course as Cr on Cr.courseId = rpc.courseId)
	inner join tmp_table_report as tmp_table on tmp_table.[key] = UD.ntid)
	where Cr.deleted=0
	and Cr.idRegion='(%RG_NAME%)'
	and Cr.idFunction (%CUSTOM_FILTER_ID%)
	and UD.SFunction (%CUSTOM_FILTER_NAME%)
	and UD.deleted=0
	and UD.Region='(%RG_NAME%)'
	order by rpc.ntid, rpc.courseId

