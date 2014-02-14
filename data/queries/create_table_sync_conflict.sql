CREATE TABLE [sync_conflict](
	[Table name] varchar(255),
	[Field name] varchar(255),
	[Field type] varchar(255),
	[Local data] varchar(255),
	[Server data] varchar(255),
	[Local timestamp] datetime,
	[Server timestamp] datetime,
	[Enable sync] yesno,
	[Row ID] varchar(255)
)