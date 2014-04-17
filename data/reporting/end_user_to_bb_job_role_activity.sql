tmp_pilot_report_4
===
	select col1 as [header], '' as [Category],'' as bColor, '' as fColor from (
		Select ActivityDetail as [col1] from activity 
		where deleted=0 
	) order by col1
===
select distinct UM.idUserdata as [value] from (user_data_mapping_role AS UM inner join user_data AS U
					on U.ntid = UM.idUserdata)
	where UM.idRegion='(%RG_NAME%)' and UM.deleted=0 and U.deleted=0 and U.Region='(%RG_NAME%)' and U.suspend=0
		and U.SFunction (%CUSTOM_FILTER_NAME%) and UM.idMapping='B'
===

select distinct ac.ActivityDetail as [Value]
from ((((user_data_mapping_role as um
inner join user_data as ud
on um.idUserdata = ud.ntid)
inner join specialism as sp
on ud.specialism = sp.SpecialismName)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join activity as ac
on spAc.idActivity = ac.id)
where um.idUserdata ='(%VALUE%)' and um.idMapping='B'
and ud.deleted=0 and um.deleted=0 and sp.deleted=0 and spAc.deleted=0
and ud.suspend=0
and ac.deleted=0 and um.idRegion='(%RG_NAME%)'

===
SELECT distinct
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
inner join tmp_pilot_report_4 as tbl_cached
on tbl_cached.[key] = UD.ntid)
Where UD.deleted=0
and UD.Region='(%RG_NAME%)'
and UD.suspend=0
and UD.SFunction (%CUSTOM_FILTER_NAME%)
ORDER BY UD.ntid
