AdHoc reporting
=======
select tmp_table_cached.NTID, tmp_table_cached.fname AS [First name], 
		tmp_table_cached.lname AS [Last name], 
		BpRoleStandard.BpRoleStandardName As [BB Job Role],
		tmp_table_cached.Sox_Indicator As [Sox Indicator], 
		tmp_table_cached.jobTitle As [Job Title] from (((
SELECT user_data.NTID, user_data.fname, 
		user_data.lname, 
		BpRoleStandard.Sox_Indicator, 
		user_data.jobTitle,
		count(user_data.NTID) AS count_ntid
FROM (((user_data INNER JOIN user_data_mapping_role ON user_data.ntid = user_data_mapping_role.idUserdata) 
			INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id) 
		INNER JOIN (
				SELECT UD.NTID,BpRoleStandard.Sox_Indicator, count(BpRoleStandard.Sox_Indicator) AS [count]
			FROM ((user_data AS UD INNER JOIN user_data_mapping_role ON UD.ntid = user_data_mapping_role.idUserdata) 
						INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id)
			WHERE user_data_mapping_role.deleted=0
			and user_data_mapping_role.idRegion='(%RG_NAME%)'
			and BpRoleStandard.deleted=0
			and UD.deleted=0
			and UD.Region='(%RG_NAME%)'
			and UD.suspend=0
			and BpRoleStandard.Sox_Indicator is not null
			and BpRoleStandard.Sox_Indicator <> ''
			and UD.SFunction (%CUSTOM_FILTER_NAME%)
			group by UD.ntid, BpRoleStandard.Sox_Indicator
	) AS tmp_table ON tmp_table.ntid = user_data.ntid)
WHERE user_data_mapping_role.deleted=0
and tmp_table.Sox_Indicator = BpRoleStandard.Sox_Indicator
and user_data_mapping_role.idRegion='(%RG_NAME%)'
and BpRoleStandard.deleted=0
and user_data.deleted=0
and user_data.Region='(%RG_NAME%)'
and user_data.suspend=0
and user_data.SFunction (%CUSTOM_FILTER_NAME%)
group by user_data.NTID, user_data.fname,user_data.lname,user_data.jobTitle,BpRoleStandard.Sox_Indicator
) as tmp_table_cached INNER JOIN user_data_mapping_role ON tmp_table_cached.ntid = user_data_mapping_role.idUserdata) 
			INNER JOIN BpRoleStandard ON user_data_mapping_role.idBpRoleStandard = BpRoleStandard.id) 
where tmp_table_cached.count_ntid > 1 and BpRoleStandard.Sox_Indicator = tmp_table_cached.Sox_Indicator
and BpRoleStandard.deleted=0 and user_data_mapping_role.deleted=0 and user_data_mapping_role.idRegion='(%RG_NAME%)'
order by tmp_table_cached.NTID, BpRoleStandard.BpRoleStandardName
=========