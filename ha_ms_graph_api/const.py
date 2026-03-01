"""Constants for the Microsoft Graph API Sandbox integration."""

DOMAIN = "ha_ms_graph_api"

# Configuration constants
CONF_CLIENT_ID = "client_id"
CONF_CLIENT_SECRET = "client_secret"
CONF_CLIENT_CERT_PATH = "client_cert_path"
CONF_TENANT_ID = "tenant_id"
CONF_UPDATE_INTERVAL = "update_interval"
CONF_SAFE_MODE = "safe_mode"
CONF_PRIVACY_MODE = "privacy_mode"
CONF_USE_CERT_AUTH = "use_cert_auth"

# Default values
DEFAULT_UPDATE_INTERVAL = 300  # seconds
DEFAULT_SAFE_MODE = True  # Enable read-only mode by default
DEFAULT_PRIVACY_MODE = True  # Hide sensitive data by default
DEFAULT_USE_CERT_AUTH = False  # Use client secret by default

# Microsoft Graph API endpoints
GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{}/oauth2/v2.0/token"
GRAPH_DEVICES_URL = "https://graph.microsoft.com/v1.0/devices"


# Device attribute keys
ATTR_DEVICE_ID = "device_id"
ATTR_DEVICE_OWNERSHIP = "device_ownership"
ATTR_ENROLLMENT_TYPE = "enrollment_type"
ATTR_IS_COMPLIANT = "is_compliant"
ATTR_OPERATING_SYSTEM = "operating_system"
ATTR_OS_VERSION = "os_version"
ATTR_MANUFACTURER = "manufacturer"
ATTR_MODEL = "model"
ATTR_LAST_SIGNIN = "last_signin"
ATTR_TRUST_TYPE = "trust_type"
ATTR_ACCOUNT_ENABLED = "account_enabled"
ATTR_DISPLAY_NAME = "display_name"
ATTR_DEVICE_GROUPS = "device_groups"
ATTR_EXTENSION_ATTRIBUTES = "extension_attributes"

# Group attribute keys
ATTR_GROUP_ID = "group_id"
ATTR_GROUP_NAME = "group_name"
ATTR_SECURITY_ENABLED = "security_enabled"
ATTR_GROUP_TYPES = "group_types"
ATTR_CREATED_DATETIME = "created_datetime"
ATTR_GROUP_MEMBERS = "group_members"

# User attribute keys
ATTR_USER_ID = "user_id"
ATTR_USER_NAME = "user_name"
ATTR_USER_MAIL = "user_mail"
ATTR_USER_PRINCIPAL_NAME = "user_principal_name"
ATTR_USER_EMPLOYEE_ID = "user_employee_id"
ATTR_USER_JOB_TITLE = "user_job_title"
ATTR_USER_DEPARTMENT = "user_department"
ATTR_USER_DEVICES = "user_devices"
