' @author Hai Lu
'
Option Compare Database

Public Const TMP_END_USER_TABLE_NAME = "tblImport"

Public Const CREATE_TABLE_END_USER_QUERY = "data\queries\create_table_user_data.sql"
Public Const DROP_TABLE_END_USER_QUERY = "data\queries\drop_table_user_data.sql"
Public Const DELETE_ALL_TABLE_END_USER_DATA_QUERY = "data\queries\delete_all_table_user_data.sql"

Public Const CREATE_TABLE_END_USER_MAPPING_QUERY = "data\queries\create_table_user_data_mapping_role.sql"

' Settings.ini
Public Const SECTION_REMOTE_DATABASE = "central database"
Public Const KEY_SERVER_NAME = "serverName"
Public Const KEY_DATABASE_NAME = "databaseName"
Public Const KEY_SYNC_TABLES = "syncTables"
Public Const KEY_PORT = "port"
Public Const KEY_USERNAME = "username"
Public Const KEY_PASSWORD = "password"

Public Const SECTION_USER_DATA = "user data"
Public Const KEY_LINE_TO_REMOVE = "linesToRemove"

Public Const SS_SYNC_TABLES = "config\synctables.ss"
Public Const SS_SYNC_USERS = "config\syncusers.ss"
' Reporting
Public Const RP_ROOT_FOLDER = "data\reporting\"
Public Const RP_END_USER_TO_SYSTEM_ROLE = "end_user_to_system_role_report"

' For tesing
Public Const END_USER_DATA_CSV_TEMPLATE_FILE_PATH = "testdata\RoleMappingDataTemplate.csv"
Public Const END_USER_DATA_CSV_TEMPLATE_TRIM_FILE_PATH = "target\RoleMappingDataTrim.csv"
Public Const END_USER_DATA_CSV_FILE_PATH = "testdata\RoleMappingData.csv"
Public Const END_USER_DATA_REPORTING_TEMPLATE = "testdata\RoleMappingNewDeploymentTemplate.xlsx"
Public Const END_USER_DATA_REPORTING_OUTPUT_DIR = "target\reporting"
Public Const END_USER_DATA_REPORTING_OUTPUT_FILE = "EndUserRoleMapping.xlsx"