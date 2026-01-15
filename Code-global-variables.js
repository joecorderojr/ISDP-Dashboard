const url = "https://docs.google.com/spreadsheets/d/1Sxcw_EVGR3VTfcWP-S3_NbBdYu1riZHmN88e2dM0ivA/edit?gid=638578937#gid=638578937";
const hrURL =
  "https://docs.google.com/spreadsheets/d/1CzfTxPZvFXgz7J0rFHxw82CR5McSCP-h5pop_YL3Rq4/edit?gid=1092268223#gid=1092268223";
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
var document_matrix_outdated = [];
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

var dataPrivacyTrainings2026 = [];
var dataPrivacyTrainingsTotal2026 = null;
var dataPrivacyTrainingsList2026 = [];
var dataPrivacyTrainings2025 = [];
var dataPrivacyTrainingsTotal2025 = null;
var dataPrivacyTrainingsList2025   = [];
var dataPrivacyTrainings2024 = [];
var dataPrivacyTrainingsTotal2024 = null;
var dataPrivacyTrainingsList2024 = [];
var cyberSecurityTrainings2026 = [];
var cyberSecurityTrainingsList2026 = [];
var cyberSecurityTrainingsTotal2026 = null;
var cyberSecurityTrainings2025 = [];
var cyberSecurityTrainingsList2025 = [];
var cyberSecurityTrainingsTotal2025 = null;
var cyberSecurityTrainings2024 = [];
var cyberSecurityTrainingsList2024 = [];
var cyberSecurityTrainingsTotal2024 = null;
var phishingResult = null;

const defaultBGColors = {
  red: 'rgba(153, 27, 27, 0.7)',
  yellow: 'rgba(161, 98, 7, 0.7)',
  green: 'rgba(21, 128, 61, 0.7)'
};

const defaultBorderColors = {
  red: 'rgba(153, 27, 27, 1)',
  yellow: 'rgba(161, 98, 7, 1)',
  green: 'rgba(21, 128, 61, 1)'
};