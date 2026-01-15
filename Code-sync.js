function updateCyberSecurity2025() {
  updateCyberSecurity("2025");
}

function updateCyberSecurity2024() {
  updateCyberSecurity("2024");
}

function updateCyberSecurity(year) {
  try {
    var ss = SpreadsheetApp.openByUrl(hrURL);
    var sheet = ss.getSheetByName("CYB101 " + year);
    var dataArray = sheet.getRange("A3:Z600").getValues();
    const data = filterEmptyRows(dataArray);

    const headers = data[0];
    const rows = data.slice(1);
    const col = headers.reduce((map, h, i) => ((map[h] = i), map), {});

    var objList = [];

    rows.forEach((row) => {
      const unit = row[col["BUSINESS UNIT"]];

      if (!unit) return;

      const inactive = row[col["INACTIVE STATUS"]];
      const status = row[col["COURSE STATUS"]];
      const score = parseFloat(row[col["QUIZ SCORE"]]) || null;
      const completionDate = row[col["COMPLETION DATE"]] || null;

      if (inactive === "ACTIVE") {
        objList.push([
          row[col["BUSINESS UNIT"]],
          row[col["COMPLETE NAME"]],
          row[col["POSITION TITLE"]],
          row[col["EMAIL ADDRESS"]],
          status,
          completionDate,
          score,
        ]);
      }
    });

    saveTrainingData(objList, "CyberSecurity" + year);

    let response = "Success";
    Logger.log(response);
    return response;
  } catch (e) {
    let response = "Failed to update file: " + e.toString();
    Logger.log(response);
    return response;
  }
}

function updateDataPrivacyTrainingData2025() {
  updateDataPrivacyTrainingData("2025");
}


function updateDataPrivacyTrainingData(year) {
  try {
    var ss = SpreadsheetApp.openByUrl(hrURL);
    var sheetName = year + " DP C/O LEGAL";
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet with name '${sheetName}' not found.`);
    }

    var lastRow = sheet.getRange("A3").getDataRegion().getLastRow();
    var data = sheet.getRange(3, 1, lastRow, 20).getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const col = headers.reduce((map, h, i) => ((map[h] = i), map), {});

    const units = {};
    var objList = [];

    rows.forEach((row) => {
      const unit = row[col["BUSINESS UNIT"]];
      if (!unit) return;

      const inactive = row[col["INACTIVE STATUS"]];
      const status = row[col["COURSE STATUS"]];
      const score = parseFloat(row[col["QUIZ SCORE"]]) || null;

      if (!units[unit]) {
        units[unit] = { employees: 0, inProgress: 0, completed: 0, scores: [] };
      }

      if (inactive === "ACTIVE") {
        objList.push([
          row[col["BUSINESS UNIT"]],
          row[col["COMPLETE NAME"]],
          row[col["POSITION TITLE"]],
          row[col["EMAIL ADDRESS"]],
          status,
          row[col["COMPLETION DATE"]],
          score,
        ]);
      }
    });

    saveTrainingData(objList, "DataPrivacy" + year);

    let response = "Success";
    Logger.log(response);
    return response;
  } catch (e) {
    let response = "Failed to update file: " + e.toString();
    Logger.log(response);
    return response;
  }
}

function saveTrainingData(list, sheetName) {
  const ss = SpreadsheetApp.openByUrl(appSettingURL);
  const sheet = ss.getSheetByName(sheetName);
  sheet.getRange("A2:G1000").clearContent();
  Logger.log("Training Data List: " + JSON.stringify(list));
  list.map(function (rowData, rowIndex) {
    sheet.getRange(rowIndex + 2, 1, 1, 7).setValues([rowData]);
  });
}
