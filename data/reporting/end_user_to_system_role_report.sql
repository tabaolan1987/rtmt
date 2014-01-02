SELECT ntid, 
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
						and sys.SystemRoleName=[SYSTEM_ROLE_NAME] 
					group by u.ntid, sys.SystemRoleName)
		AS [Procurement Catalogue Approver],
			(SELECT TOP 1 (IIF(NOT ISNULL(bpRole.BpRoleStandardName),"Y",""))
					FROM ((user_data_mapping_role AS userMapping
						INNER JOIN user_data AS u
							ON u.ntid = userMapping.idUserdata)
						INNER JOIN BpRoleStandard AS bpRole
							ON bpRole.idBpRoleStandard = userMapping.idBpRoleStandard)
					WHERE bpRole.BpRoleStandardName = [BP_ROLE_STANDARD_NAME] 
						AND ntid=UD.ntid  
					GROUP BY ntid, bpRole.BpRoleStandardName)
		AS [POQR Approver]
FROM	user_data AS UD