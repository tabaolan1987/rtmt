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
	[siteLocation] AS [Maximo Site Location],
	[purchasOrg] AS [Purchasing Org],
	[Optionalfield1] AS [Optional Field 1],
	[Optionalfield2] AS [Optional Field 2]
FROM user_data ORDER BY [ntid]
=====
SELECT {% 
				SELECT SR.SystemRoleName FROM SystemRole AS SR 
														INNER JOIN SystemRoleCategory AS SRC 
														ON SR.idSystemRoleCategory= SRC.idSystemRoleCategory 
													ORDER BY SRC.SystemRoleCategory, SR.SystemRoleName 
			| 
				(SELECT TOP 1 (IIF(NOT ISNULL(sys.SystemRoleName), "X","")) 
					FROM (((((user_data_mapping_role AS userMapping
						INNER JOIN user_data AS u
							ON u.ntid = userMapping.idUserdata)
						INNER JOIN BpRoleStandard AS bpRole
							ON bpRole.idBpRoleStandard = userMapping.idBpRoleStandard)
						INNER JOIN MappingSystemRoleBpStandardRole AS mappSysBp
							ON mappSysBp.idBpRoleStandard = userMapping.idBpRoleStandard)
						INNER JOIN SystemRole AS sys
							ON  mappSysBp.idSystemRole = sys.idSystemRole)
						INNER JOIN MappingRegionSystemRole AS mappRe
							ON userMapping.idRegion = mappRe.idRegion)
					where mappRe.idSystemRole = mappSysBp.idSystemRole
						and u.ntid = UD.ntid 
						and sys.SystemRoleName='(%VALUE%)'
					group by u.ntid, sys.SystemRoleName)
				AS [(%VALUE%)] 
		%}
FROM user_data AS UD ORDER BY UD.ntid
=====
SELECT 
		{% 
				SELECT BRS.BpRoleStandardName
					FROM BpRoleStandard AS BRS 
						INNER JOIN BpRoleStandardCategory AS BRSC 
							ON BRS.idBpRoleStandardCategory=BRSC.idBpRoleStandardCategory 
					ORDER BY BRSC.BpRoleStandardCategoryName, BRS.BpRoleStandardName 
			|
				(SELECT TOP 1 (IIF(NOT ISNULL(bpRole.BpRoleStandardName),"Y",""))
					FROM ((user_data_mapping_role AS userMapping
						INNER JOIN user_data AS u
							ON u.ntid = userMapping.idUserdata)
						INNER JOIN BpRoleStandard AS bpRole
							ON bpRole.idBpRoleStandard = userMapping.idBpRoleStandard)
					WHERE bpRole.BpRoleStandardName = '(%VALUE%)'
						AND ntid=UD.ntid  
					GROUP BY ntid, bpRole.BpRoleStandardName)
				AS [(%VALUE%)] 
		%}
FROM user_data AS UD ORDER BY UD.ntid