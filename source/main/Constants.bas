' @author Hai Lu
'
Option Compare Database

Public Const TMP_END_USER_TABLE_NAME = "tblImport"

Public Const CREATE_TABLE_END_USER_QUERY = "data\queries\create_table_user_data.sql"
Public Const DROP_TABLE_END_USER_QUERY = "data\queries\drop_table_user_data.sql"
Public Const DELETE_ALL_TABLE_END_USER_DATA_QUERY = "data\queries\delete_all_table_user_data.sql"

' Settings.ini
Public Const SECTION_REMOTE_DATABASE = "central database"
Public Const KEY_SERVER_NAME = "serverName"
Public Const KEY_DATABASE_NAME = "databaseName"
Public Const KEY_SYNC_TABLES = "syncTables"
Public Const KEY_PORT = "port"
Public Const KEY_USERNAME = "username"
Public Const KEY_PASSWORD = "password"

Public Const SS_SYNC_TABLES = "config\synctables.ss"
                    
' For tesing
Public Const END_USER_DATA_CSV_FILE_PATH = "testdata\RoleMappingData.csv"
Public Const END_USER_DATA_REPORTING_TEMPLATE = "testdata\RoleMappingNewDeploymentTemplate.xlsx"
Public Const END_USER_DATA_REPORTING_OUTPUT_DIR = "target\reporting"
Public Const END_USER_DATA_REPORTING_OUTPUT_FILE = "EndUserRoleMapping.xlsx"