function updateCyberSecurity2025() {
  try 
  {
    var ss = SpreadsheetApp.openByUrl(cyberSecurityURL);
    var sheet = ss.getSheetByName("CYB101");
    var dataArray = sheet.getRange("A3:Z600").getValues();
    const data = filterEmptyRows(dataArray);
    
    const headers = data[0];
    const rows = data.slice(1);
    const col = headers.reduce((map, h, i) => (map[h] = i, map), {});

    var objList = [];

    rows.forEach(row => {
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
          score
        ]);
      }
    });

    saveCyberSecurity(objList, "CyberSecurity2025");
    
    let response = "Success";
    Logger.log(response);
    return response;
  }
  catch (e) {
    let response = "Failed to update file: " + e.toString();
    Logger.log(response);
    return response;
  }
}


function updateCyberSecurity2024() 
{
  try 
  {
    var ss = SpreadsheetApp.openByUrl(cyberSecurityURL);
    var sheet = ss.getSheetByName("CYB101 2024");
    var dataArray = sheet.getRange("A3:Z600").getValues();
    const data = filterEmptyRows(dataArray);
    
    const headers = data[0];
    const rows = data.slice(1);

    const col = headers.reduce((map, h, i) => (map[h] = i, map), {});
    
    var objList = [];

    rows.forEach(row => {
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
          score
        ]);
      }
    });

    saveCyberSecurity(objList, "CyberSecurity2024");

    let response = "Success";
    Logger.log(response);
    return response;
  }
  catch (e) {
    let response = "Failed to update file: " + e.toString();
    Logger.log(response);
    return response;
  }
}

function saveCyberSecurity(list, sheetName) {
  const ss = SpreadsheetApp.openByUrl(appSettingURL);
  const sheet = ss.getSheetByName(sheetName);
  sheet.getRange("A2:G1000").clearContent();

  list.map(function (rowData, rowIndex) {
    sheet.getRange(rowIndex + 2, 1, 1, 7).setValues([rowData]);
  });
}
