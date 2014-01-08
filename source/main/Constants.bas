' @author Hai Lu
'
Option Compare Database

Public Const TMP_END_USER_TABLE_NAME = "tblImport"
'
Public Const END_USER_DATA_TABLE_NAME = "user_data"

' Queries
Public Const QUERIES_DIR = "data\queries\"
Public Const Q_CREATE = 1
Public Const Q_INSERT = 2
Public Const Q_UPDATE = 3
Public Const Q_DELETE_ALL = 4
Public Const Q_CUSTOM = 0

' Settings.ini
Public Const SETTINGS_FILE = "data\config\settings.ini"
Public Const SECTION_REMOTE_DATABASE = "central database"
Public Const KEY_SERVER_NAME = "serverName"
Public Const KEY_DATABASE_NAME = "databaseName"
Public Const KEY_SYNC_TABLES = "syncTables"
Public Const KEY_PORT = "port"
Public Const KEY_USERNAME = "username"
Public Const KEY_PASSWORD = "password"

Public Const SECTION_USER_DATA = "user data"
Public Const KEY_LINE_TO_REMOVE = "linesToRemove"
Public Const KEY_TABLE_NAME = "tableNames"
Public Const KEY_REGION_NAME = "regionName"

Public Const SECTION_APPLICATION = "application"
Public Const KEY_LOG_LEVEL = "logLevel"

' System setting file
Public Const SS_DIR = "data\config\"
Public Const SS_SYNC_TABLES = "synctables"
Public Const SS_SYNC_USERS = "syncusers"

' Reporting
Public Const RP_SPLIT_LEVEL_1 = "====="
Public Const RP_SPLIT_LEVEL_2 = "==="
Public Const RP_SECTION_TYPE_FIXED = "fixed"
Public Const RP_SECTION_TYPE_AUTO = "auto"
Public Const RP_CONFIG_FILE_EXTENSION = ".cfg"
Public Const RP_QUERY_FILE_EXTENSION = ".sql"
Public Const RP_TEMPLATE_FILE_EXTENSION = ".xlsx"
Public Const RP_ROOT_FOLDER = "data\reporting\"
Public Const RP_END_USER_TO_SYSTEM_ROLE = "end_user_to_system_role_report"

' For tesing
Public Const END_USER_DATA_CSV_TEMPLATE_FILE_PATH = "testdata\RoleMappingDataTemplate.csv"
Public Const END_USER_DATA_CSV_TEMPLATE_TRIM_FILE_PATH = "target\RoleMappingDataTrim.csv"
Public Const END_USER_DATA_CSV_FILE_PATH = "testdata\RoleMappingData.csv"
Public Const END_USER_DATA_REPORTING_TEMPLATE = "testdata\RoleMappingNewDeploymentTemplate.xlsx"
Public Const END_USER_DATA_REPORTING_OUTPUT_DIR = "target\reporting"
Public Const END_USER_DATA_REPORTING_OUTPUT_FILE = "EndUserRoleMapping.xlsx"

Public Const END_USER_DATA_FILE_XLSX = "testdata\EndUserRoleMapping.xlsx"
Public Const END_USER_DATA_FILE_CSV = "target\EndUserRoleMapping.csv"