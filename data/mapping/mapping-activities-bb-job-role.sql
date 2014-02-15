select [id], [BpRoleStandardName] As [VALUE], '' As [COMMENT] from BpRoleStandard WHERE deleted=0 order by [BpRoleStandardName]
=====
select id, 
		(ActivityDetail + DetailPlus) as [VALUE], 
		('- Group: ' + ActivityGroup + CHR(13) + CHR(10) + '- Process: ' + ActivityProcess + CHR(13) + CHR(10) + '- Training: ' + ActivityTraining) as [COMMENT]
		from Activity where deleted=0 order by ActivityDetail
=====
select Description As [MappingChar],deleted from MappingActivityBpStandardRole
	where idActivity='(%ID_TOP%)' and  idBpRoleStandard='(%ID_LEFT%)' and function_region='(%RG_F_ID%)'
=====
update MappingActivityBpStandardRole set deleted='(%CHECK%)', Description='(%VALUE%)'
	where idActivity='(%ID_TOP%)' and  idBpRoleStandard='(%ID_LEFT%)' and function_region='(%RG_F_ID%)'
=====
insert into MappingActivityBpStandardRole(id, idActivity, idBpRoleStandard,Description, function_region, deleted) 
	values('(%ID%)', '(%ID_TOP%)', '(%ID_LEFT%)','(%VALUE%)', '(%RG_F_ID%)', '0')