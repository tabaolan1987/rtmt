' @author Hai Lu
'
Option Compare Database
Public Const SKIP_LAUNCH = "skip.launch"

Public Const TMP_END_USER_TABLE_NAME = "tblImport"

Public Const END_USER_DATA_TABLE_NAME = "user_data"
'
Public Const END_USER_DATA_CACHE_TABLE_NAME = "user_data_cache"

' Help content
Public Const HELP_FIX_READONLY_PERMISSION = "data\helps\fix-readonly-permission.html"

Public Const HELP_ERROR_EUDL = "data\helps\error-eudl.txt"

Public Const HELP_UPLOAD_EUDL = "data\helps\upload-eudl.txt"

Public Const HELP_REPORTS = "data\helps\reports.txt"

Public Const HELP_ADMIN = "data\helps\admin.txt"

Public Const HELP_MAPPING = "data\helps\mapping.txt"

Public Const HELP_EDIT_EUDL = "data\helps\edit-eudl.txt"

Public Const TEMPLATE_EUDL = "data\template\eudl_template.xlsx"

Public Const FIELD_TIMESTAMP = "Timestamp"
Public Const FIELD_ID = "id"
Public Const FIELD_DELETED = "Deleted"
Public Const FIELD_FIRST_NAME = "fname"
Public Const FIELD_LAST_NAME = "lname"
Public Const FIELD_SELECT = "Select"
Public Const FIELD_DB_FIELD = "Db field"
Public Const FIELD_SPECIALISM = "specialism"
Public Const FIELD_REGION_FUNCTION = "SFunction"
Public Const FIELD_MAPPING_CHAR = "MappingChar"

Public Const TABLE_SYNC_CONFLICT = "sync_conflict"
Public Const TABLE_USER_DATA_CONFLICT = "user_data_conflict"
Public Const TABLE_USER_DATA_DUPLICATE = "user_data_duplicate"
Public Const TABLE_USER_DATA_LDAP_CONFLICT = "user_data_ldap_conflict"
Public Const TABLE_USER_DATA_LDAP_NOTFOUND = "user_data_ldap_notfound"
Public Const TABLE_TMP_TABLE_REPORT = "tmp_table_report"
Public Const TABLE_USER_PRIVILEGES = "user_privileges"
Public Const TABLE_MAPPING_SPECIALISM_ACITIVITY = "mapping_specialism_acitivity"
Public Const TABLE_USER_DATA_MAPPING_ROLE = "user_data_mapping_role"
Public Const TABLE_AUDIT_LOG = "audit_logs"

Public Const TEXT_DEFAULT_SELECT_REGION = "Select Region"
Public Const TEXT_DEFAULT_SELECT_FUNCTION = "Select Function"
Public Const TEXT_DEFAULT_SELECT_MAPPING_TYPE = "Select mapping type"

' Environment
Public Const ENV_DEVELOP = "DEVELOP"
Public Const ENV_DEV = "DEV"
Public Const ENV_INT = "INT"
Public Const ENV_PROD = "PROD"

'Role & Permission
Public Const P_R_ADMIN = "Admin"
Public Const P_R_RC = "RC"
Public Const P_R_TC = "TC"
Public Const P_R_AM = "AM"

Public Const P_W = "W"
Public Const P_R = "R"

' Queries
Public Const QUERIES_DIR = "data\queries\"
Public Const Q_CREATE = 1
Public Const Q_INSERT = 2
Public Const Q_UPDATE = 3
Public Const Q_DELETE_ALL = 4
Public Const Q_SELECT = 5
Public Const Q_CUSTOM = 0

Public Const Q_KEY_VALUE = "VALUE"
Public Const Q_KEY_CHECK = "CHECK"
Public Const Q_KEY_COMMENT = "COMMENT"
Public Const Q_KEY_ID = "ID"
Public Const Q_KEY_ID_TOP = "ID_TOP"
Public Const Q_KEY_ID_LEFT = "ID_LEFT"
Public Const Q_KEY_FUNCTION_REGION_ID = "RG_F_ID"
Public Const Q_KEY_REGION_NAME = "RG_NAME"
Public Const Q_KEY_FUNCTION_REGION_NAME = "RG_F_NAME"
Public Const Q_KEY_MAPPING_FIELDS = "MAPPING_FIELDS"
Public Const Q_KEY_FILTER = "FILTER"

' for mapping section
Public Const Q_TOP = 5
Public Const Q_LEFT = 6
Public Const Q_CHECK = 7

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
Public Const KEY_TEST_NTID = "testNtid"
Public Const KEY_CHECK_IP_URL = "checkIpURL"

Public Const KEY_VALIDATOR_URL = "validatorUrl"
Public Const KEY_TOKEN = "token"
Public Const KEY_BULK_SIZE = "bulkSize"
Public Const KEY_NTID_FIELD = "ntidField"

Public Const SECTION_APPLICATION = "application"
Public Const KEY_VERSION = "version"
Public Const KEY_ENV = "env"
Public Const KEY_LOG_LEVEL = "logLevel"
Public Const KEY_ENABLE_TESTING = "enableTesting"
Public Const KEY_ENABLE_VALIDATION = "enableValidation"
Public Const KEY_SYNC_FIXED_TABLES = "syncFixedTables"
Public Const KEY_ENABLE_AUDITLOG = "enableAuditLog"

Public Const SECTION_GENERAL = "General"
Public Const KEY_NAME = "name"
Public Const KEY_QUERY_TYPE = "queryType"
Public Const KEY_WORK_SHEET = "workSheet"
Public Const KEY_PIVOT_TABLE = "pivotTable"
Public Const KEY_PIVOT_TABLE_NAME = "pivotTableName"
Public Const KEY_PIVOT_WORD_WRAP_COLS = "pivotWordWrapCols"
Public Const KEY_PIVOT_TABLE_WORK_SHEET = "pivotWorkSheet"
Public Const SECTION_FORMAT = "Format"
Public Const KEY_FILL_HEADER = "fillHeader"
Public Const KEY_START_HEADER_ROW = "startHeaderRow"
Public Const KEY_START_HEADER_COL = "startHeaderCol"
Public Const KEY_FILL_CATEGORY = "fillCategory"
Public Const KEY_START_CATEGORY_ROW = "startCategoryRow"
Public Const KEY_START_ROW = "startRow"
Public Const KEY_START_COL = "startCol"
Public Const KEY_SKIP_CHECK_HEADER = "skipCheckHeader"

Public Const SECTION_TOP = "Top"
Public Const SECTION_LEFT = "Left"

Public Const KEY_MERGE_ENABLE = "mergeEnable"
Public Const KEY_MERGE_COLUMES = "mergeCols"
Public Const KEY_MERGE_PRIMARY = "mergePrimary"
Public Const KEY_CUSTOM_MODE = "customMode"

Public Const KEY_MAPPING_CHAR = "mappingChar"

' System setting file
Public Const SS_DIR = "data\config\"
Public Const SS_SYNC_TABLES = "synctables"
Public Const SS_SYNC_ROLE_TABLES = "syncroletables"
Public Const SS_SYNC_MAPPING_TABLES = "syncmappingtables"
Public Const SS_SYNC_USERS = "syncusers"
Public Const SS_VALIDATOR_MAPPING = "validatormapping"
Public Const SS_JUNK_TABLES = "junktables"


Public Const SYNC_TYPE_DEFAULT = 0
Public Const SYNC_TYPE_ROLE = 1
Public Const SYNC_TYPE_MAPPING = 2

' File
Public Const FILE_EXTENSION_CONFIG = ".ini"
Public Const FILE_EXTENSION_QUERY = ".sql"
Public Const FILE_EXTENSION_QUERY_MAPPING = ".sqlm"
Public Const FILE_EXTENSION_TEMPLATE = ".xlsx"
Public Const FILE_EXTENSION_REPORT = ".xlsx"
Public Const SPLIT_LEVEL_1 = "====="
Public Const SPLIT_LEVEL_2 = "==="

' Mapping
Public Const MAPPING_ACTIVITIES_SPECIALISM = "mapping-activities-specialism"
Public Const MAPPING_ACTIVITIES_BB_JOB_ROLE = "mapping-activities-bb-job-role"
Public Const MAPPING_ROOT_FOLDER = "data\mapping\"

' Reporting
Public Const RP_TYPE_DEFAULT = "default"
Public Const RP_TYPE_MAPPING = "mapping"

Public Const RP_QUERY_TYPE_SECTION = "section"
Public Const RP_QUERY_TYPE_SIMPLE = "simple"

Public Const RP_SECTION_TYPE_FIXED = "fixed"
Public Const RP_SECTION_TYPE_AUTO = "auto"
Public Const RP_SECTION_TYPE_TMP_TABLE = "tmp_table"
Public Const RP_SECTION_TYPE_TMP_PILOT_REPORT = "tmp_pilot_report"

Public Const RP_ROOT_FOLDER = "data\reporting\"
Public Const RP_END_USER_TO_SYSTEM_ROLE = "end_user_to_system_role_report"
Public Const RP_END_USER_TO_BB_JOB_ROLE = "end_user_to_bb_job_role_report"
Public Const RP_ROLE_MAPPING_OUTPUT_OF_TOOL_FOR_SECURITY = "role_mapping_output_of_tool_for_security"
Public Const RP_END_USER_TO_COURSE = "end_user_to_course_report"
Public Const RP_AUDIT_LOG = "Audit_Log_Report"

Public Const RP_DEFAULT_OUTPUT_FOLDER = "target\reporting"

' For tesing
Public Const END_USER_DATA_CSV_TEMPLATE_FILE_PATH = "testdata\RoleMappingDataTemplate.csv"
Public Const END_USER_DATA_CSV_TEMPLATE_TRIM_FILE_PATH = "target\RoleMappingDataTrim.csv"
Public Const END_USER_DATA_CSV_FILE_PATH = "testdata\RoleMappingData.csv"
Public Const END_USER_DATA_REPORTING_TEMPLATE = "testdata\RoleMappingNewDeploymentTemplate.xlsx"
Public Const END_USER_DATA_REPORTING_OUTPUT_DIR = "target\reporting"
Public Const END_USER_DATA_REPORTING_OUTPUT_FILE = "EndUserRoleMapping.xlsx"

Public Const END_USER_DATA_FILE_XLSX = "testdata\eudl_template.xlsx"
Public Const END_USER_DATA_FILE_CSV = "target\EndUserRoleMapping.csv"
Public Const END_USER_DATA_TMP_FILE_CSV = "user_data_tmp.csv"
Public Const END_USER_DATA_TMP_FINAL_FILE_CSV = "user_data_tmp_final.csv"