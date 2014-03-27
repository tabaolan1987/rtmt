tmp_pilot_report_1
===
	select col1 as [header], col2 as [Category], bColor, fColor from (
	select B.BpRoleStandardName as [col1], IIF(ISNULL(BG.BpRoleStandardCategoryName), "MASTER DATA",BG.BpRoleStandardCategoryName) As [col2]
		, BG.orderPriority, BG.bColor, BG.fColor
	 from (BpRoleStandard AS B left join BpRoleStandardCategory AS BG on BG.id = B.idBpRoleStandardCategory)
	where B.deleted = 0 and BG.deleted = 0
	) order by orderPriority, col1
===
select idUserdata as [value] from user_data_mapping_role where idRegion='(%RG_NAME%)' and deleted=0
===
select bpRole.BpRoleStandardName AS [value]
from (user_data_mapping_role as UMR 
inner join BpRoleStandard as bpRole
on UMR.idBpRoleStandard = bpRole.id)
where UMR.idUserdata = '(%VALUE%)' and UMR.Deleted=0
and bpRole.Deleted = 0
and UMR.idRegion='(%RG_NAME%)'
===
SELECT 
	UD.[NTID],
	[fname] AS [First Name],
	[lname] AS [Last Name],
	[jobTitle] AS [Job Title],
	'' AS [Performace Unit],
	[purchasingOrg] AS [Purchasing Org],
	[siteLocation] AS [Maximo Site Location],
	(%MAPPING_FIELDS%)
FROM (user_data as UD
left join tmp_pilot_report_1 as tbl_cached
on tbl_cached.[key] = UD.ntid)
Where UD.deleted=0
and UD.Region='(%RG_NAME%)'
ORDER BY UD.ntid