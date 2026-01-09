function getLoginUser() {
  loginUser = Session.getActiveUser().getEmail();
  getAccessRole();

  Logger.log("loginUser: " + loginUser);
}

function getAccessRole()
{
    var ss = SpreadsheetApp.openByUrl(appSettingURL);
    var sheet = ss.getSheetByName("Settings");

    var roleDocuments = sheet.getRange("A2:A500").getValues();
    roleDocuments = filterEmptyRows(roleDocuments).flat();

    var roleTrainings = sheet.getRange("B2:B500").getValues();
    roleTrainings = filterEmptyRows(roleTrainings).flat();

    hasAccessRoleInTrainings = roleTrainings.includes(loginUser);
    hasAccessRoleInDocuments = roleDocuments.includes(loginUser);

    Logger.log(roleDocuments);
    Logger.log(roleTrainings);
    Logger.log("hasAccessRoleInTrainings: " + hasAccessRoleInTrainings);
    Logger.log("hasAccessRoleInDocuments: " + hasAccessRoleInDocuments);
}
