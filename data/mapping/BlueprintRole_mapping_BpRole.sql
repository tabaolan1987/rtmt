select [id], [BpRoleStandardName] As [VALUE], '' As [COMMENT] from BpRoleStandard WHERE deleted=0 order by [BpRoleStandardName]
=====
select id, [BPrintName] as [VALUE], '' As [COMMENT] from BlueprintRoles where deleted=0 order by [BPrintName]
=====
select * from BlueprintRole_mapping_BpRole where idBpRole = '(%ID_LEFT%)' and idBluePrintRole = '(%ID_TOP%)'
=====
update BlueprintRole_mapping_BpRole set deleted='(%CHECK%)'
	where idBpRole='(%ID_LEFT%)' and  idBluePrintRole='(%ID_TOP%)'
=====
insert into BlueprintRole_mapping_BpRole(id, idBluePrintRole, idBpRole, deleted) 
	values('(%ID%)', '(%ID_TOP%)', '(%ID_LEFT%)', '0')