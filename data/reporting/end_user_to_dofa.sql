End user to DofA (PO)
=======
SELECT user_data.NTID, user_data.fname AS [First name], user_data.lname AS [Last name], BpRoleStandard.BpRoleStandardName AS [BB Job Role], user_data.jobTitle As [Job Title]
FROM ((user_data INNER JOIN user_data_mapping_role ON user_data.ntid = user_data_mapping_role.idUserdata) 
			INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id) 
		LEFT JOIN Dofa ON user_data.ntid = Dofa.username1
WHERE BpRoleStandard.Dofa_Type='PO' AND Dofa.[DOA_Spend_Limit] Is Null
and user_data_mapping_role.deleted=0
and user_data_mapping_role.idRegion='(%RG_NAME%)'
and Dofa.deleted=0
and BpRoleStandard.deleted=0
and user_data.deleted=0
and user_data.Region='(%RG_NAME%)'
and user_data.suspend=0
and user_data.SFunction (%CUSTOM_FILTER_NAME%)
ORDER BY user_data.ntid, BpRoleStandard.BpRoleStandardName
=========
End user to DofA (IN)
=======
SELECT user_data.NTID, user_data.fname AS [First name], user_data.lname AS [Last name], BpRoleStandard.BpRoleStandardName AS [BB Job Role], user_data.jobTitle As [Job Title]
FROM ((user_data INNER JOIN user_data_mapping_role ON user_data.ntid = user_data_mapping_role.idUserdata) 
			INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id) 
		LEFT JOIN Dofa ON user_data.ntid = Dofa.username1
WHERE BpRoleStandard.Dofa_Type='IN' AND Dofa.[DOA_Spend_Limit] Is Null
and user_data_mapping_role.deleted=0
and user_data_mapping_role.idRegion='(%RG_NAME%)'
and Dofa.deleted=0
and BpRoleStandard.deleted=0
and user_data.deleted=0
and user_data.Region='(%RG_NAME%)'
and user_data.suspend=0
and user_data.SFunction (%CUSTOM_FILTER_NAME%)
ORDER BY user_data.ntid, BpRoleStandard.BpRoleStandardName
