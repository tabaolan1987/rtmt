EUDL (Raw)
=======
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
	[mappingTypeBpRoles] AS [Mapping type],
	[mapped_bb_job_roles] AS [Backbone Job Roles],
	[mapped_qualifications] AS [Qualifications],
	[day1user] AS [Day1 User],
	[courserq] AS [Course Requirements],
	[NTR],
	[Timestamp],
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
	[Spare35] AS [Optional Field 35]
FROM user_data as UD
Where UD.deleted=0
and UD.suspend=0
and UD.Region='(%RG_NAME%)'
and UD.SFunction (%CUSTOM_FILTER_NAME%)
ORDER BY UD.ntid
=========
EUDL to Activity (Raw)
=======
select ud.NTID, ud.fname AS [First name], 
	ud.lname AS [Last name], ud.mapped_bb_job_roles AS [BB Job Roles],
	ud.[mappingTypeBpRoles] AS [Mapping type],
	ud.jobTitle As [Job Title],
	ac.Column1 As [Column 1],
	ac.ActivityName As [Activity Name],
	ac.ActivityGroup As [Activity Group],
	ac.ActivityTraining As [Activity Training],
	ac.ActivityDetail As [Activity Detail]
from ((((user_data_mapping_role as um
inner join user_data as ud
on um.idUserdata = ud.ntid)
inner join specialism as sp
on ud.specialism = sp.SpecialismName)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join activity as ac
on spAc.idActivity = ac.id)
where um.idMapping='B'
and ud.deleted=0 and um.deleted=0 and sp.deleted=0 and spAc.deleted=0
and ud.suspend=0
and ac.deleted=0 and um.idRegion='(%RG_NAME%)'
and ud.SFunction (%CUSTOM_FILTER_NAME%)
order by ud.ntid
=========
EUDL to DofA (Raw)
=======
SELECT user_data.NTID, user_data.fname AS [First name], 
	user_data.lname AS [Last name], 
	user_data.jobTitle As [Job Title],
	BpRoleStandard.BpRoleStandardName AS [BB Job Role], 
	user_data.[mappingTypeBpRoles] AS [Mapping type],
	BpRoleStandard.DofA_Type As [DofA type],
	BpRoleStandard.Sox_Indicator As [Sox Indicator],
	Dofa.sno As [Dofa SNO],
	Dofa.DOA_SRM_Au As [DOA SRM Au],
	Dofa.Employee_G As [Employee G],
	Dofa.username2 As [User name],
	Dofa.DOA_Spend_Limit As [DOA Spend Limit],
	Dofa.Crcy,
	Dofa.changeOn As [Changed on],
	Dofa.timechange As [Time],
	Dofa.changeby As [Changed by]
FROM (((user_data INNER JOIN user_data_mapping_role ON user_data.ntid = user_data_mapping_role.idUserdata) 
			INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id) 
		LEFT JOIN Dofa ON user_data.ntid = Dofa.username1)
WHERE  user_data_mapping_role.deleted=0
and user_data_mapping_role.idRegion='(%RG_NAME%)'
and ((Dofa.deleted IS NULL) OR Dofa.deleted=0)
and BpRoleStandard.deleted=0
and ((Dofa.[DOA_SRM_Au] IS NULL) or BpRoleStandard.Dofa_Type = Dofa.[DOA_SRM_Au])
and user_data.deleted=0
and user_data.Region='(%RG_NAME%)'
and user_data.suspend=0
and user_data.SFunction (%CUSTOM_FILTER_NAME%)
ORDER BY user_data.ntid, BpRoleStandard.BpRoleStandardName