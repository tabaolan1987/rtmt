
create table Region(
	idRegion int IDENTITY(1,1) NOT NULL,
	RegionName varchar(200)
)

create table SystemRoleCategory(
	idSystemRoleCategory int IDENTITY(1,1) NOT NULL,
	SystemRoleCategory varchar(500)
)

create table SystemRole(
	idSystemRole int identity(1,1) not null,
	idSystemRoleCategory int,
	SystemRoleName varchar(500)
)

create table BpRoleStandardCategory(
	idBpRoleStandardCategory int identity(1,1) not null,
	BpRoleStandardCategoryName varchar(500)
)

create table BpRoleStandard(
	idBpRoleStandard int identity(1,1) not null,
	idBpRoleStandardCategory int,
	BpRoleStandardName varchar(500)
)
create table Activity(
	idActivity int identity(1,1) not null,
	ActivityName varchar(500),
	ActivityComment varchar(500),
	ActivityDofa int
)

create table MappingRegionSystemRole(
	idMappingRegionSystemRole int identity(1,1) not null,
	idRegion int,
	idSystemRole int
)

create table MappingSystemRoleBpStandardRole(
	idMappingSystemRoleBpStandardRole int identity(1,1) not null,
	idSystemRole int,
	idBpRoleStandard int
)

create table MappingActivityBpStandardRole(
	idMappingActivityBpStandardRole int identity(1,1) not null,
	idActivity int,
	idBpRoleStandard int
)
