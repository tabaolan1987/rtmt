-- @author: Hai Lu
--
-- Update database from v0.2.2 to v0.2.3
-- Change:
--	+ Alter table [audit_logs] 
--			increase column [prev_value] & [new_value] size from 255 to 4000
--	+ Alter table [user_data]
--			add new column [actor_ntid], [mapped_bb_job_roles], [mapped_qualifications]
--	+ New tables [user_action_mapping], [user_change_log], [user_data_tmp]	
--	+ New triggers [trg_create_user_change_log], [trg_update_user_change_log]
--		for Audit user data change
--	

-- SELECT DATABASE 
USE [upstream_role_mapping]
--				 --
-- ALTER TABLE --
--				 --
GO
ALTER TABLE [audit_logs]
ALTER COLUMN [prev_value] VARCHAR(4000)
GO
ALTER TABLE [audit_logs]
ALTER COLUMN [new_value] VARCHAR(4000)
GO
ALTER TABLE [user_data]
ADD [actor_ntid] VARCHAR(255)
GO
ALTER TABLE [user_data]
ADD [mapped_bb_job_roles] VARCHAR(4000)
GO
ALTER TABLE [user_data]
ADD [mapped_qualifications] VARCHAR(4000)
--				--
-- UPDATE DATA  --
--	  			--
GO
UPDATE user_data SET actor_ntid = ''
GO
DECLARE @ntid VARCHAR(255)
DECLARE @region VARCHAR(255)
DECLARE db_cursor CURSOR FOR  
				SELECT [ntid], [region]
				FROM [user_data] WHERE deleted=0
DECLARE @qualifications VARCHAR(4000)
DECLARE @roles VARCHAR(4000)
OPEN db_cursor 
FETCH NEXT FROM db_cursor INTO @ntid, @region
WHILE @@FETCH_STATUS = 0   
BEGIN 
	SET @qualifications = ''
	SELECT @qualifications = @qualifications + Qname + ', ' FROM (SELECT DISTINCT q.Qname
											FROM (user_data_mapping_qualification
											AS um inner join Qualifications AS q
											ON q.id = um.idQuali)
											WHERE um.ntid=@ntid AND q.deleted=0 AND um.deleted=0
											AND um.idRegion=@region) AS cached_table
											ORDER BY Qname
	SET @qualifications = SUBSTRING(@qualifications, 0, LEN(@qualifications))
	SET @roles = ''
	SELECT @roles = @roles + BpRoleStandardName + ', ' FROM (select DISTINCT bpRole.BpRoleStandardName
											from (user_data_mapping_role as UMR 
											inner join BpRoleStandard as bpRole
											on UMR.idBpRoleStandard = bpRole.id)
											where UMR.idUserdata = @ntid and UMR.Deleted=0
											and bpRole.Deleted = 0
											and UMR.idRegion=@region) AS cached_table
											ORDER BY BpRoleStandardName
	SET @roles = SUBSTRING(@roles, 0, LEN(@roles))
	UPDATE user_data SET mapped_qualifications = @qualifications, mapped_bb_job_roles = @roles
			WHERE ntid=@ntid AND region = @region
	FETCH NEXT FROM db_cursor INTO @ntid, @region
END
CLOSE db_cursor   
DEALLOCATE db_cursor
--				 --
-- CREATE TABLE  --
--				 --
GO
CREATE TABLE [user_action_mapping]
(
	[data_field] VARCHAR(255) PRIMARY KEY,
	[action] VARCHAR(255)
)
GO
CREATE TABLE [user_change_log]
(
	[id] BIGINT IDENTITY PRIMARY KEY,
	[action] VARCHAR(255),
	[eu_ntid] VARCHAR(255),
	[eu_first_name] VARCHAR(255),
	[eu_last_name] VARCHAR(255),
	[data_fields] VARCHAR(255),
	[prev_value] VARCHAR(4000),
	[new_value] VARCHAR(4000),
	[table_name] VARCHAR(255),
	[description] VARCHAR(4000),
	[actor_ntid] VARCHAR(255),
	[region] VARCHAR(255),
	[Timestamp] DATETIME DEFAULT (getdate()),
	[deleted] BIT DEFAULT ((0))
)
GO
CREATE TABLE [user_data_tmp]
(
	[id] VARCHAR(100),
	[ntid] VARCHAR(255),
	[gpid] VARCHAR(255),
	[fname] VARCHAR(255),
	[lname] VARCHAR(255),
	[email] VARCHAR(255),
	[omsSubfunction] VARCHAR(255),
	[departmentBusiness] VARCHAR(255),
	[specialism] VARCHAR(255),
	[jobTitle] VARCHAR(255),
	[sponsorForeName] VARCHAR(255),
	[sponsorSurname] VARCHAR(255),
	[vta] VARCHAR(255),
	[country] VARCHAR(255),
	[contractor] VARCHAR(100),
	[changeNetworkLevel] VARCHAR(255),
	[dofa] VARCHAR(100),
	[dofaType] VARCHAR(100),
	[region] VARCHAR(100),
	[SFunction] VARCHAR(255),
	[SdSubFunction] VARCHAR(255),
	[siteLocation] VARCHAR(255),
	[purchasingOrg] VARCHAR(255),
	[STeam] VARCHAR(255),
	[blueprintRole] VARCHAR(255),
	[suspend] BIT,
	[sponsorNTID] VARCHAR(255),
	[actor_ntid] VARCHAR(100),
	[mappingTypeBpRoles] VARCHAR(255),
	[mapped_bb_job_roles] VARCHAR(4000),
	[mapped_qualifications] VARCHAR(4000),
	[Timestamp] DATETIME,
	[Deleted] BIT,
	[spare1] VARCHAR(255),
	[spare2] VARCHAR(255),
	[spare3] VARCHAR(255),
	[spare4] VARCHAR(255),
	[spare5] VARCHAR(255),
	[spare6] VARCHAR(255),
	[spare7] VARCHAR(255),
	[spare8] VARCHAR(255),
	[spare9] VARCHAR(255),
	[spare10] VARCHAR(255),
	[spare11] VARCHAR(255),
	[spare12] VARCHAR(255),
	[spare13] VARCHAR(255),
	[spare14] VARCHAR(255),
	[spare15] VARCHAR(255),
	[spare16] VARCHAR(255),
	[spare17] VARCHAR(255),
	[spare18] VARCHAR(255),
	[spare19] VARCHAR(255),
	[spare20] VARCHAR(255),
	[spare21] VARCHAR(255),
	[spare22] VARCHAR(255),
	[spare23] VARCHAR(255),
	[spare24] VARCHAR(255),
	[spare25] VARCHAR(255),
	[spare26] VARCHAR(255),
	[spare27] VARCHAR(255),
	[spare28] VARCHAR(255),
	[spare29] VARCHAR(255),
	[spare30] VARCHAR(255),
	[spare31] VARCHAR(255),
	[spare32] VARCHAR(255),
	[spare33] VARCHAR(255),
	[spare34] VARCHAR(255),
	[spare35] VARCHAR(255),
	[table_type] VARCHAR(255)
)
--					 --
-- CREATE PROCEDURE  --
--					 --
GO
--			   	  --
-- CREATE TRIGGER --
--			   	  --
GO
CREATE TRIGGER [trg_create_user_change_log] ON [user_data]
AFTER INSERT
AS
BEGIN
	PRINT 'Detect insert new EUDL record'
	INSERT INTO [user_change_log]([action], [eu_ntid], 
								[eu_first_name], [eu_last_name], 
								[table_name],[actor_ntid], 
								[region]) 
			SELECT 'Create central store record' AS [action],
					[ntid] AS [eu_ntid],
					[fname] AS [eu_first_name],
					[lname] AS [eu_last_name],
					'user_data' AS [table_name],
					[actor_ntid],
					[region]
			FROM inserted
END
GO
CREATE TRIGGER [trg_update_user_change_log] ON [user_data]
AFTER UPDATE
AS
BEGIN
	DECLARE @rid VARCHAR(255)
	DECLARE @actor_ntid VARCHAR(255)
	DECLARE @eu_ntid VARCHAR(255)
	DECLARE @eu_first_name VARCHAR(255)
	DECLARE @eu_last_name VARCHAR(255)
	DECLARE @region VARCHAR(255)
	DECLARE @new_value VARCHAR(4000)
	DECLARE @prev_value VARCHAR(4000)
	DECLARE @data_field VARCHAR(255)
	DECLARE @action VARCHAR(255)
	DECLARE @count int
	DECLARE @retvalOUT VARCHAR(4000)
	DECLARE @query NVARCHAR(1000)
	-- Only trigger if only one record update
	-- 
	SET @count = (SELECT COUNT(*) FROM inserted)
	IF @count = 1
	BEGIN
		-- Ignore error
		SET XACT_ABORT OFF;
		PRINT 'Detect update. Count:' + CONVERT(VARCHAR(255), @count)
		BEGIN TRY
			DECLARE db_cursor CURSOR FOR  
				SELECT [data_field], [action]
				FROM [user_action_mapping]
			PRINT 'Prepare data'
			SET @rid = (SELECT [id] FROM inserted)
			SET @actor_ntid = (SELECT [actor_ntid] FROM inserted)
			SET @eu_ntid = (SELECT [ntid] FROM inserted)
			SET @eu_first_name = (SELECT [fname] FROM inserted)
			SET @eu_last_name = (SELECT [lname] FROM inserted)
			SET @region = (SELECT [region] FROM inserted)
			PRINT 'Insert cached data'
			INSERT [user_data_tmp]
				([id],[ntid], [gpid],[fname],[lname],[email],[omsSubfunction],[departmentBusiness],
					[specialism],[jobTitle],[sponsorForeName],[sponsorSurname],[vta],[country],[contractor],
					[changeNetworkLevel],[dofa],[dofaType],[region],[SFunction],[SdSubFunction],[siteLocation],
					[purchasingOrg],[STeam],[blueprintRole],[suspend],[sponsorNTID],[actor_ntid],
					[mappingTypeBpRoles],[mapped_bb_job_roles],[mapped_qualifications],[Timestamp],[Deleted],
					[spare1],[spare2],[spare3],[spare4],[spare5],[spare6],[spare7],[spare8],[spare9],[spare10],
					[spare11],[spare12],[spare13],[spare14],[spare15],[spare16],[spare17],[spare18],[spare19],[spare20],
					[spare21],[spare22],[spare23],[spare24],[spare25],[spare26],[spare27],[spare28],[spare29],[spare30],
					[spare31],[spare32],[spare33],[spare34],[spare35],
					[table_type])
				SELECT
				[id],[ntid], [gpid],[fname],[lname],[email],[omsSubfunction],[departmentBusiness],
					[specialism],[jobTitle],[sponsorForeName],[sponsorSurname],[vta],[country],[contractor],
					[changeNetworkLevel],[dofa],[dofaType],[region],[SFunction],[SdSubFunction],[siteLocation],
					[purchasingOrg],[STeam],[blueprintRole],[suspend],[sponsorNTID],[actor_ntid],
					[mappingTypeBpRoles],[mapped_bb_job_roles],[mapped_qualifications],[Timestamp],[Deleted],
					[spare1],[spare2],[spare3],[spare4],[spare5],[spare6],[spare7],[spare8],[spare9],[spare10],
					[spare11],[spare12],[spare13],[spare14],[spare15],[spare16],[spare17],[spare18],[spare19],[spare20],
					[spare21],[spare22],[spare23],[spare24],[spare25],[spare26],[spare27],[spare28],[spare29],[spare30],
					[spare31],[spare32],[spare33],[spare34],[spare35],
					'inserted' AS [table_type]
				FROM [inserted]
			INSERT [user_data_tmp]
				([id],[ntid], [gpid],[fname],[lname],[email],[omsSubfunction],[departmentBusiness],
					[specialism],[jobTitle],[sponsorForeName],[sponsorSurname],[vta],[country],[contractor],
					[changeNetworkLevel],[dofa],[dofaType],[region],[SFunction],[SdSubFunction],[siteLocation],
					[purchasingOrg],[STeam],[blueprintRole],[suspend],[sponsorNTID],[actor_ntid],
					[mappingTypeBpRoles],[mapped_bb_job_roles],[mapped_qualifications],[Timestamp],[Deleted],
					[spare1],[spare2],[spare3],[spare4],[spare5],[spare6],[spare7],[spare8],[spare9],[spare10],
					[spare11],[spare12],[spare13],[spare14],[spare15],[spare16],[spare17],[spare18],[spare19],[spare20],
					[spare21],[spare22],[spare23],[spare24],[spare25],[spare26],[spare27],[spare28],[spare29],[spare30],
					[spare31],[spare32],[spare33],[spare34],[spare35],
					[table_type])
				SELECT
				[id],[ntid], [gpid],[fname],[lname],[email],[omsSubfunction],[departmentBusiness],
					[specialism],[jobTitle],[sponsorForeName],[sponsorSurname],[vta],[country],[contractor],
					[changeNetworkLevel],[dofa],[dofaType],[region],[SFunction],[SdSubFunction],[siteLocation],
					[purchasingOrg],[STeam],[blueprintRole],[suspend],[sponsorNTID],[actor_ntid],
					[mappingTypeBpRoles],[mapped_bb_job_roles],[mapped_qualifications],[Timestamp],[Deleted],
					[spare1],[spare2],[spare3],[spare4],[spare5],[spare6],[spare7],[spare8],[spare9],[spare10],
					[spare11],[spare12],[spare13],[spare14],[spare15],[spare16],[spare17],[spare18],[spare19],[spare20],
					[spare21],[spare22],[spare23],[spare24],[spare25],[spare26],[spare27],[spare28],[spare29],[spare30],
					[spare31],[spare32],[spare33],[spare34],[spare35],
					'deleted' AS [table_type]
				FROM [deleted]
			PRINT 'Open cursor'
			OPEN db_cursor 
			PRINT 'Loop field mapping'  
			FETCH NEXT FROM db_cursor INTO @data_field, @action
			WHILE @@FETCH_STATUS = 0   
			BEGIN 
				PRINT 'Check data field: ' + @data_field + '. Action: ' + @action
				SET @query = N'SELECT @valOUT = [' + @data_field + N'] FROM [user_data_tmp] where [table_type]=''inserted'' and [id]=''' + @rid + N''''
				EXEC sp_executesql @query, N'@valOUT VARCHAR(4000) OUTPUT', @valOUT=@new_value OUTPUT;
				SET @query = N'SELECT @valOUT = [' + @data_field + N'] FROM [user_data_tmp] where [table_type]=''deleted'' and [id]=''' + @rid + N''''
				EXEC sp_executesql @query, N'@valOUT VARCHAR(4000) OUTPUT', @valOUT=@prev_value OUTPUT;
				PRINT 'Prev value: ' + @prev_value + '. New value: ' + @new_value
				IF NOT @new_value = @prev_value
				BEGIN
					PRINT 'Detect ' + @action + '. From: ' + @prev_value + ' to: ' + @new_value
					INSERT INTO [user_change_log]([action], [eu_ntid], 
											[eu_first_name], [eu_last_name], 
											[table_name],[actor_ntid],
											[data_fields],
											[prev_value], [new_value],
											[region])
						VALUES(@action, @eu_ntid, @eu_first_name, @eu_last_name,
									'user_data', @actor_ntid, @data_field, 
									@prev_value, @new_value, @region)
				END
				FETCH NEXT FROM db_cursor INTO @data_field, @action
			END
			PRINT 'Close cursor'
			CLOSE db_cursor   
			DEALLOCATE db_cursor
			-- Check mapping BB Job Roles
			DECLARE @mapping_type varchar(255)
			DECLARE @prev_mapping_type varchar(255)
			SET @mapping_type = (SELECT [mappingTypeBpRoles] FROM inserted)
			SET @prev_mapping_type = (SELECT [mappingTypeBpRoles] FROM deleted)
			
			PRINT 'Prev mapping: ' + @prev_mapping_type + '. New mapping: ' + @mapping_type
			IF NOT @prev_mapping_type = @mapping_type
			BEGIN
				PRINT 'Detect change mapping type. From: ' + @prev_mapping_type + ' to: ' + @mapping_type
				INSERT INTO [user_change_log]([action], [eu_ntid], 
										[eu_first_name], [eu_last_name], 
										[table_name],[actor_ntid],
										[data_fields],
										[prev_value], [new_value],
										[region])
					VALUES('Change mapping type', @eu_ntid, @eu_first_name, @eu_last_name,
								'user_data', @actor_ntid,'mappingTypeBpRoles', 
								@prev_mapping_type, @mapping_type, @region)
			END
			DECLARE @new_mapping VARCHAR(4000)
			DECLARE @prev_mapping VARCHAR(4000)
			SET @prev_mapping = (SELECT [mapped_bb_job_roles] FROM deleted)
			SET @new_mapping = (SELECT [mapped_bb_job_roles] FROM inserted)
			PRINT 'Check mapping roles: ' + @new_mapping
			PRINT 'Prev roles: ' + @prev_mapping + '. New roles: ' + @new_mapping
			IF NOT @prev_mapping = @new_mapping
			BEGIN
				PRINT 'Detect change mapping role. From: ' + @prev_mapping + ' to: ' + @new_mapping
				INSERT INTO [user_change_log]([action], [eu_ntid], 
										[eu_first_name], [eu_last_name], 
										[table_name],[actor_ntid],
										[data_fields],
										[prev_value], [new_value],
										[region])
					VALUES('Update Backbone Job Roles. Mapping type: ' + @mapping_type, @eu_ntid, @eu_first_name, @eu_last_name,
								'user_data', @actor_ntid,'mapped_bb_job_roles', 
								@prev_mapping, @new_mapping, @region)
			END
			PRINT 'Delete cached data'
			DELETE FROM [user_data_tmp] WHERE [id] = @rid
		END TRY
		BEGIN CATCH
			PRINT 'Could not create change log'
		END CATCH
	END
	ELSE
		PRINT 'Detect multi records. Count: ' + CONVERT(VARCHAR(255), @count)
END
GO
--					--
-- INSERT RAW DATA  --
--	  			    --
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('gpid','Update GPID')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('fname','Update First name')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('lname','Update Last name')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('email','Update Email')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('omsSubfunction','Update Function (OMS)/ Sub-function')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('departmentBusiness','Update Department or Business Unit')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('specialism','Update Specialism')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('jobTitle','Update Job Title')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('sponsorForeName','Update Line Manager/ Sponsor Forename')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('sponsorSurname','Update Line Manager/ Sponsor Surname')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('vta','Update VTA')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('country','Update Country')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('contractor','Update Contractor? (Y/N)')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('changeNetworkLevel','Update Change Network Level')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('dofa','Update Dofa')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('dofaType','Update Dofa Type')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('siteLocation','Update Maximo Site Location')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('purchasingOrg','Update Purchasing Org')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('SFunction','Update Standard Function')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('SdSubFunction','Update Standard Sub Function')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('STeam','Update Standard Team')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('blueprintRole','Update Blueprint Role')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('suspend','Update Suspened status')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('sponsorNTID','Update Line Manager/ Sponsor NTID')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('mapped_qualifications','Change Qualification mapping')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare1','Update Optional Field 1')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare2','Update Optional Field 2')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare3','Update Optional Field 3')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare4','Update Optional Field 4')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare5','Update Optional Field 5')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare6','Update Optional Field 6')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare7','Update Optional Field 7')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare8','Update Optional Field 8')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare9','Update Optional Field 9')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare10','Update Optional Field 10')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare11','Update Optional Field 11')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare12','Update Optional Field 12')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare13','Update Optional Field 13')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare14','Update Optional Field 14')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare15','Update Optional Field 15')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare16','Update Optional Field 16')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare17','Update Optional Field 17')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare18','Update Optional Field 18')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare19','Update Optional Field 19')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare20','Update Optional Field 20')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare21','Update Optional Field 21')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare22','Update Optional Field 22')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare23','Update Optional Field 23')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare25','Update Optional Field 24')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare25','Update Optional Field 25')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare26','Update Optional Field 26')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare27','Update Optional Field 27')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare28','Update Optional Field 28')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare29','Update Optional Field 29')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare30','Update Optional Field 30')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare31','Update Optional Field 31')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare32','Update Optional Field 32')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare33','Update Optional Field 33')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare34','Update Optional Field 34')
GO
INSERT [user_action_mapping]([data_field], [action]) VALUES('spare35','Update Optional Field 35')