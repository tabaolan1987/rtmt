tmp_pilot_report
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
	[GPID],
	[fname] AS [First Name],
	[lname] AS [Last Name],
	[email] AS [E-mail address],
	[omsSubfunction] AS [Function (OMS)/ Sub-function],
	[departmentBusiness] AS [Department or Business Unit],
	[Specialism],
	[jobTitle] AS [Job Title],
	[sponsorForeName] AS [Line Manager/ Sponsor Forename],
	[sponsorSurname] AS [Line Manager/ Sponsor Surname],
	[VTA],
	[Country],
	[contractor] AS [Contractor?(Y/N)],
	[SFunction] AS [Standard Function],
	[SdSubFunction] AS [Standard Sub Function],
	[STeam] AS [Standard Team],
	[blueprintRole] AS [Blueprint Role],
	[Region],
	[sponsorNTID] AS [Line Manager/ Sponsor NTID],
	[purchasingOrg] AS [Purchasing Org],
	[siteLocation] AS [Maximo Site Location],
	[Spare1] AS [Optional Field 1],
	[Spare2] AS [Optional Field 2],
	[Spare3] AS [Optional Field 3],
	[Spare4] AS [Optional Field 4],
	[Spare5] AS [Optional Field 5],
	[Spare6] AS [Optional Field 6],
	[Spare7] AS [Optional Field 7],
	[Spare8] AS [Optional Field 8],
	[Spare9] AS [Optional Field 9],
	[Spare10] AS [Optional Field 10],
	[Spare11] AS [Optional Field 11],
	[Spare12] AS [Optional Field 12],
	[Spare13] AS [Optional Field 13],
	[Spare14] AS [Optional Field 14],
	[Spare15] AS [Optional Field 15],
	[Spare16] AS [Optional Field 16],
	[Spare17] AS [Optional Field 17],
	[Spare18] AS [Optional Field 18],
	[Spare19] AS [Optional Field 19],
	[Spare20] AS [Optional Field 20],
	[Spare21] AS [Optional Field 21],
	[Spare22] AS [Optional Field 22],
	[Spare23] AS [Optional Field 23],
	[Spare24] AS [Optional Field 24],
	[Spare25] AS [Optional Field 25],
	[Spare26] AS [Optional Field 26],
	[Spare27] AS [Optional Field 27],
	[Spare28] AS [Optional Field 28],
	[Spare29] AS [Optional Field 29],
	[Spare30] AS [Optional Field 30],
	[Spare31] AS [Optional Field 31],
	[Spare32] AS [Optional Field 32],
	[Spare33] AS [Optional Field 33],
	[Spare34] AS [Optional Field 34],
	[Spare35] AS [Optional Field 35],
	(%MAPPING_FIELDS%)
FROM (user_data as UD
left join tmp_pilot_report as tbl_cached
on tbl_cached.[key] = UD.ntid)
Where UD.deleted=0
and UD.Region='(%RG_NAME%)'
ORDER BY UD.ntid