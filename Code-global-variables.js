const url = "https://docs.google.com/spreadsheets/d/1Sxcw_EVGR3VTfcWP-S3_NbBdYu1riZHmN88e2dM0ivA/edit?gid=638578937#gid=638578937";
const dataPrivacyURL =
  "https://docs.google.com/spreadsheets/d/1CzfTxPZvFXgz7J0rFHxw82CR5McSCP-h5pop_YL3Rq4/edit?gid=56851791#gid=56851791";
const cyberSecurityURL = "https://docs.google.com/spreadsheets/d/1fmfKouNUT1tLwOxIdMQoZEvks2C8DHoK6TqT7N_Alcc/edit?gid=2024740077#gid=2024740077"
const appSettingURL = "https://docs.google.com/spreadsheets/d/1WvLg2sRLzP8xaROWkBtcfgFzyZJhawQXP_O1PxqD0ao/edit?gid=1428041287#gid=1428041287";

var loginUser = ""; 
var document_list = [];
var document_types = [];
var unit_list = [];
var data_doc_total = [];
var document_statuses = [];
var document_matrix_raw = [];
var document_matrix = [];
var document_matrix_forReview = [];
var document_matrix_overdue = [];
var document_matrix_reviewed = [];
var tools_data;
var tools_matrix = [];
var npc_checklist_data;
var iso_27001_data;
var vaSummary_data;
var soc2Type2_data;
var iso_27001_compliance_data = []; 
var hasAccessRoleInTrainings = false;
var hasAccessRoleInDocuments = false;
var hasAccessRoleInSync = false;
var vaPeriod = "";

var dataPrivacyTrainings2025 = [];
var dataPrivacyTrainingsTotal2025 = null;
var dataPrivacyTrainingsList2025   = [];
var cyberSecurityTrainings2025 = [];
var cyberSecurityTrainingsList2025 = [];
var cyberSecurityTrainingsTotal2025 = null;
var cyberSecurityTrainings2024 = [];
var cyberSecurityTrainingsList2024 = [];
var cyberSecurityTrainingsTotal2024 = null;
var phishingResult = null;