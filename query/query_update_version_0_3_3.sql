-- @author: Hai Lu
--
-- Update database from v0.3 to v0.3.3
-- to store EUDL's uploaded date
-- Change:
--	+ Alter table [user_data]
--			add new column [ext_timestamp] 
--	+ Alter table [user_data_tmp]
--			add new column [ext_timestamp] 
--	+ Alter trigger trg_update_user_change_log
--  + Update ext_timestamp by old timestamp (without hours,minutes,seconds)

-- SELECT DATABASE 
USE [upstream_role_mapping]
--				 --
-- ALTER TABLE --
--				 --
GO
ALTER TABLE [user_data]
ADD [ext_timestamp] datetime
GO
ALTER TABLE [user_data_tmp]
ADD [ext_timestamp] datetime

--			   	  --
-- ALTER TRIGGER --
--			   	  --
GO
ALTER TRIGGER [trg_update_user_change_log] ON [user_data]
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
					[table_type],[ext_timestamp])
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
					'inserted' AS [table_type], [ext_timestamp]
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
					[table_type],[ext_timestamp])
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
					'deleted' AS [table_type],[ext_timestamp]
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

--				--
-- UPDATE DATA  --
--	  			--
GO
UPDATE [user_data] SET [ext_timestamp] = cast(floor(cast([timestamp] as float)) as datetime) WHERE [deleted] = 0
