' @author Hai Lu
'
Option Compare Database

Public Const TMP_END_USER_TABLE_NAME = "tblImport"

Public Const END_USER_DATA_TABLE_NAME = "user_data"
'
Public Const END_USER_DATA_CACHE_TABLE_NAME = "user_data_cache"

Public Const FIELD_TIMESTAMP = "Timestamp"
Public Const FIELD_ID = "id"
Public Const FIELD_DELETED = "Deleted"
Public Const FIELD_FIRST_NAME = "fname"
Public Const FIELD_LAST_NAME = "lname"

Public Const TABLE_SYNC_CONFLICT = "sync_conflict"
Public Const TABLE_USER_DATA_CONFLICT = "user_data_conflict"
Public Const TABLE_USER_DATA_DUPLICATE = "user_data_duplicate"
' Queries
Public Const QUERIES_DIR = "data\queries\"
Public Const Q_CREATE = 1
Public Const Q_INSERT = 2
Public Const Q_UPDATE = 3
Public Const Q_DELETE_ALL = 4
Public Const Q_CUSTOM = 0

' settings ini
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
Public Const KEY_REGION_FUNCTION_ID = "regionFunctionId"

Public Const KEY_VALIDATOR_URL = "validatorUrl"
Public Const KEY_TOKEN = "token"
Public Const KEY_BULK_SIZE = "bulkSize"
Public Const KEY_NTID_FIELD = "ntidField"

Public Const SECTION_APPLICATION = "application"
Public Const KEY_LOG_LEVEL = "logLevel"

Public Const SECTION_GENERAL = "General"
Public Const KEY_NAME = "name"
Public Const KEY_QUERY_TYPE = "queryType"
Public Const KEY_WORK_SHEET = "workSheet"
Public Const SECTION_FORMAT = "Format"
Public Const KEY_FILL_HEADER = "fillHeader"
Public Const KEY_START_HEADER_ROW = "startHeaderRow"
Public Const KEY_START_HEADER_COL = "startHeaderCol"
Public Const KEY_START_ROW = "startRow"
Public Const KEY_START_COL = "startCol"

' System setting file
Public Const SS_DIR = "data\config\"
Public Const SS_SYNC_TABLES = "synctables"
Public Const SS_SYNC_USERS = "syncusers"
Public Const SS_VALIDATOR_MAPPING = "validatormapping"

' Reporting
Public Const RP_QUERY_TYPE_SECTION = "section"
Public Const RP_QUERY_TYPE_SIMPLE = "simple"
Public Const RP_SPLIT_LEVEL_1 = "====="
Public Const RP_SPLIT_LEVEL_2 = "==="
Public Const RP_SECTION_TYPE_FIXED = "fixed"
Public Const RP_SECTION_TYPE_AUTO = "auto"
Public Const RP_CONFIG_FILE_EXTENSION = ".ini"
Public Const RP_QUERY_FILE_EXTENSION = ".sql"
Public Const RP_TEMPLATE_FILE_EXTENSION = ".xlsx"
Public Const RP_REPORT_FILE_EXTENSION = ".xlsx"
Public Const RP_ROOT_FOLDER = "data\reporting\"
Public Const RP_END_USER_TO_SYSTEM_ROLE = "end_user_to_system_role_report"
Public Const RP_DEFAULT_OUTPUT_FOLDER = "target\reporting"

' For tesing
Public Const END_USER_DATA_CSV_TEMPLATE_FILE_PATH = "testdata\RoleMappingDataTemplate.csv"
Public Const END_USER_DATA_CSV_TEMPLATE_TRIM_FILE_PATH = "target\RoleMappingDataTrim.csv"
Public Const END_USER_DATA_CSV_FILE_PATH = "testdata\RoleMappingData.csv"
Public Const END_USER_DATA_REPORTING_TEMPLATE = "testdata\RoleMappingNewDeploymentTemplate.xlsx"
Public Const END_USER_DATA_REPORTING_OUTPUT_DIR = "target\reporting"
Public Const END_USER_DATA_REPORTING_OUTPUT_FILE = "EndUserRoleMapping.xlsx"

Public Const END_USER_DATA_FILE_XLSX = "testdata\EndUserRoleMapping.xls"
Public Const END_USER_DATA_FILE_CSV = "target\EndUserRoleMapping.csv"
Public Const END_USER_DATA_TMP_FILE_CSV = "user_data_tmp.csv"
Public Const END_USER_DATA_TMP_FINAL_FILE_CSV = "user_data_tmp_final.csv"