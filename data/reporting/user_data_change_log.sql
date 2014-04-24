User data changelog
=======
select top 60000 timestamp, action, eu_ntid, eu_first_name, eu_last_name,
	prev_value, new_value, table_name, description, actor_ntid, region
	 from user_change_log where deleted=0 and region='(%RG_NAME%)'
	 order by timestamp desc
=========