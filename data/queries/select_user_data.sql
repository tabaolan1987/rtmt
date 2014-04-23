SELECT 
	*
FROM user_data as UD
Where UD.deleted=0
and UD.Region='(%RG_NAME%)'
and UD.suspend=0