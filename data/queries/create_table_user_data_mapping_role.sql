CREATE TABLE [user_data_mapping_role](
	[id]  varchar(50) PRIMARY KEY,
	[idUserdata] varchar(255),
	[idRegion] varchar(50),
	[idBpRoleStandard] varchar(50),
	[Timestamp] datetime,
	[Deleted] bit
)