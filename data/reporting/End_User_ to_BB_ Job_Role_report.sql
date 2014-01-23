SELECT 
	[NTID],
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
	[contractor] AS [Contractor?],
	[changeNetworkLevel] AS [Change Network Level],
	[dofa] AS [DofA (Y/N)],
	[dofaType] AS [DofA Type  (Indent, Commitment, Sourcing)],
	[siteLocation] AS[Maximo Site Location],
	[purchasingOrg] AS [Purchasing Org],
	[changeNetworkLevel] AS [Change Network Level 2],
	[Spare1],
	[Spare2],
	[Spare3],
	[Spare4],
	[Spare5],
	[Spare6],
	[Spare7],
	[Spare8],
	[Spare9],
	[Spare10],
	[Spare11],
	[Spare12],
	[Spare13],
	[Spare14],
	[Spare15],
	[Spare16],
	[Spare17],
	[Spare18],
	[Spare19],
	[Spare20]
FROM user_data ORDER BY [ntid]
=====
SELECT 
	{% Central Desktop Confirmation Requester,Contract Owner
		| 
		(select top 1 IIF(NOT ISNULL(bpRole.BpRoleStandardName), "Yes","") 
from (((((user_data as udata 
inner join specialism as sp
on sp.SpecialismName = udata.specialism)
inner join SpecialismMappingActivity as spAc
on sp.id = spAc.idSpecialism)
inner join Activity as ac
on spAc.idActivity =ac.idActivity)
inner join MappingActivityBpStandardRole as AcBpMapp
on ac.idActivity = AcBpMapp.idActivity)
inner join BpRoleStandard as bpRole
on AcBpMapp.idBpRoleStandard = bpRole.id)
where udata.ntid = UD.ntid
and bpRole.BpRoleStandardName = '(%VALUE%)'
and spAc.function_region='(%RG_F_ID%)')
				AS [(%VALUE%)]
			%}
FROM user_data AS UD ORDER BY UD.ntid

