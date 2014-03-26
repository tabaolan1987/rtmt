select id, [SpecialismName] as [VALUE], [SpecialismDescription] As [COMMENT] from specialism where deleted=0 order by [SpecialismName]
=====
select [id], [BpRoleStandardName] As [VALUE], '' As [COMMENT] from BpRoleStandard WHERE deleted=0 order by [BpRoleStandardName]
=====
select * from specialism_mapping_BpRole where idSpecialism = '(%ID_LEFT%)' and idBpRole = '(%ID_TOP%)'
=====
update specialism_mapping_BpRole set deleted='(%CHECK%)'
	where idBpRole='(%ID_TOP%)' and  idSpecialism='(%ID_LEFT%)'
=====
insert into specialism_mapping_BpRole(id, idBpRole, idSpecialism, deleted) 
	values('(%ID%)', '(%ID_TOP%)', '(%ID_LEFT%)', '0')