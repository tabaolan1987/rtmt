fixed
===
NTID
GPID
First Name
Last Name
E-mail address
Function (OMS)/ Sub-function
Department or Business Unit
Specialism
Job Title
Line Manager/ Sponsor Forename
Line Manager/ Sponsor Surname
VTA
Country
Contractor?
Change Network Level
DofA (Y/N)
DofA Type  (Indent, Commitment, Sourcing)
Maximo Site Location
Purchasing Org
Optional Field 1
Optional Field 2
===
SELECT 
	[ntid],
	[gpid],
	[fname],
	[lname],
	[email],
	[omsSubfunction],
	[departmentBusiness],
	[specialism],
	[jobTitle],
	[sponsorForeName],
	[sponsorSurname],
	[vta],
	[country],
	[contractor],
	[changeNetworkLevel],
	[dofa],
	[dofaType],
	[siteLocation],
	[purchasOrg],
	[Optionalfield1],
	[Optionalfield2]
FROM user_data ORDER BY [ntid]
=====
auto
===
SELECT SR.SystemRoleName AS [VAL_OUT] FROM SystemRole AS SR 
														INNER JOIN SystemRoleCategory AS SRC 
														ON SR.idSystemRoleCategory= SRC.idSystemRoleCategory 
													ORDER BY SRC.SystemRoleCategory, SR.SystemRoleName
===
SELECT {% 
				SELECT SR.SystemRoleName AS [VAL_OUT] FROM SystemRole AS SR 
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
						and sys.SystemRoleName='(%VAL_IN%)'
					group by u.ntid, sys.SystemRoleName)
				AS [(%VAL_COL%)] 
		%}
FROM user_data AS UD ORDER BY UD.ntid
=====
auto
===
SELECT BRS.BpRoleStandardName AS [VAL_OUT] 
					FROM BpRoleStandard AS BRS 
						INNER JOIN BpRoleStandardCategory AS BRSC 
							ON BRS.idBpRoleStandardCategory=BRSC.idBpRoleStandardCategory 
					ORDER BY BRSC.BpRoleStandardCategoryName, BRS.BpRoleStandardName
===
SELECT 
		{% 
				SELECT BRS.BpRoleStandardName AS [VAL_OUT] 
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
					WHERE bpRole.BpRoleStandardName = '(%VAL_IN%)'
						AND ntid=UD.ntid  
					GROUP BY ntid, bpRole.BpRoleStandardName)
				AS [(%VAL_COL%)] 
		%}
FROM user_data AS UD ORDER BY UD.ntid