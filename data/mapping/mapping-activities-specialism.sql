select id, [SpecialismName] as [VALUE], [SpecialismDescription] As [COMMENT] from specialism where deleted='0'
=====
select id, 
		([ActivityDetail] + [DetailPlus]) as [VALUE], 
		('Group: ' + [ActivityGroup] + CHR(13) + CHR(10) + 'Process: ' + [ActivityProcess] + CHR(13) + CHR(10) + 'Training' + [ActivityTraining]) as [COMMENT]
		from Activity where deleted='0'
=====
select * from SpecialismMappingActivity 
	where idActivity='(%ID_TOP%)' and  idSpecialism='(%ID_LEFT%)' and function_region='(%RG_F_ID%)'
=====
update SpecialismMappingActivity set deleted='(%CHECK%)'
	where idActivity='(%ID_TOP%)' and  idSpecialism='(%ID_LEFT%)' and function_region='(%RG_F_ID%)'
=====
insert into SpecialismMappingActivity(id, idActivity, idSpecialism, function_region, deleted) 
	values('(%ID%)', '(%ID_TOP%)', '(%ID_LEFT%)', '(%RG_F_ID%)', '0')