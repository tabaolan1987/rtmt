select id, [SpecialismName] as [VALUE], [SpecialismDescription] As [COMMENT] from specialism where deleted=0 and [SpecialismName] in ((%FILTER%)) order by [SpecialismName]
=====
select id, 
		(ActivityDetail + DetailPlus) as [VALUE], 
		('- Group: ' + ActivityGroup + CHR(13) + CHR(10) + '- Process: ' + ActivityProcess + CHR(13) + CHR(10) + '- Training: ' + ActivityTraining) as [COMMENT]
		from Activity where deleted=0 order by ActivityDetail
=====
select * from SpecialismMappingActivity 
	where idActivity='(%ID_TOP%)' and  idSpecialism='(%ID_LEFT%)' and idRegion='(%RG_NAME%)'
=====
update SpecialismMappingActivity set deleted='(%CHECK%)'
	where idActivity='(%ID_TOP%)' and  idSpecialism='(%ID_LEFT%)' and idRegion='(%RG_NAME%)'
=====
insert into SpecialismMappingActivity(id, idActivity, idSpecialism, idRegion, deleted) 
	values('(%ID%)', '(%ID_TOP%)', '(%ID_LEFT%)', '(%RG_NAME%)', '0')