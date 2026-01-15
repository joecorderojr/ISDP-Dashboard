function getNPCChecklist() {
  try {
    var ss = SpreadsheetApp.openByUrl(appSettingURL);
    var sheet = ss.getSheetByName("NPC Requirements Checklist");
    const lastRow = sheet.getLastRow();

    const data = sheet.getRange(2, 2, lastRow - 1).getValues();

    let yesCount = 0;
    let noCount = 0;

    data.forEach(row => {
      if (row[0] === "Yes") yesCount++;
      if (row[0] === "No") noCount++;
    });
    
    npc_checklist_data = {
      labels: ['Complied', 'Not Complied'],
      values: [yesCount, noCount],
    };
  } catch (e) {
    Logger.log('Error fetching NPC Checklist data: ' + e.toString());
    return null;
  }
}

function getISO27001Data() {
  try {
    const ss = SpreadsheetApp.openByUrl(appSettingURL);
    const sheet = ss.getSheetByName('ISO27001');
    const data = sheet.getDataRange().getValues();

    const headers = data.shift();
    const unitIndex = headers.indexOf("Business Unit");
    const statusIndex = headers.indexOf("Status");

    const summary = {};

    data.forEach(row => {
      const unit = row[unitIndex];
      const status = row[statusIndex];

      if (!unit || !status) return;

      if (!summary[unit]) {
        summary[unit] = {
          forReview: 0,
          outdated: 0,
          revised: 0
        };
      }

      switch (status) {
        case "For Review":
          summary[unit].forReview++;
          break;
        case "Outdated":
          summary[unit].outdated++;
          break;
        case "Revised":
          summary[unit].revised++;
          break;
      }
    });
    const units = Object.keys(summary);
    const forReview = units.map(u => summary[u].forReview || 0);
    const outdated = units.map(u => summary[u].outdated || 0);
    const revised = units.map(u => summary[u].revised || 0);
    iso_27001_data = {
      labels: units,
      forReview,
      outdated,
      revised
    };
  } catch (e) {
    Logger.log('Error fetching ISO270001 data: ' + e.toString());
    return null;
  }
}

function getSOC2TYPE2ChartData() {
  try {
    const ss = SpreadsheetApp.openByUrl(appSettingURL);
    const sheet = ss.getSheetByName('SOC2 TYPE2');
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    const labels = [];
    const values = [];

    data.forEach(row => {
      labels.push(row[0]);
      values.push(Math.round(row[3] * 100));
    });

    soc2Type2_data = {
      labels,
      values,
    }
  } catch (e) {
    Logger.log('Error fetching SOC2 TYPE2 data: ' + e.toString());
    return null;
  }
}

function getDashboardData() {
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName("Updated Masterlist");

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const unitCol = headers.indexOf("Unit") + 1;
  const statusCol = headers.indexOf("For Review Status") + 1;

  var lastrow = sheet.getRange("A1").getDataRegion().getLastRow();
  var data = sheet.getRange(2,1,lastrow,14).getValues();

  const col = {};
  let totalReview = 0;
  let totalOutdated = 0;
  let totalReviewed = 0;

  data.forEach(row => {
    const unit = row[unitCol - 1];
    const status = row[statusCol - 1];

    if (!col[unit]) {
      col[unit] = { "For Review": 0, "Outdated": 0, "Reviewed": 0 };
    }

    col[unit][status]++;
    if (status === "For Review") totalReview++;
    if (status === "Outdated") totalOutdated++;
    if (status === "Reviewed") totalReviewed++;
  });

  const tableRows = [];
  let grandTotal = totalReview + totalOutdated + totalReviewed;

  for (const unit in col) {
    const forReview = col[unit]["For Review"];
    const outdated = col[unit]["Outdated"];
    const reviewed = col[unit]["Reviewed"];
    const total = forReview + outdated + reviewed;

    tableRows.push({
      unit,
      forReview,
      outdated,
      reviewed,
      total,
      pctforReview: total ? forReview / total : 0,
      pctOutdated: total ? outdated / total : 0,
      pctReviewed: total ? reviewed / total : 0
    });
  }

  // Summary TOTAL row (top row)
  tableRows.push({
    unit: "Total",
    forReview: totalReview,
    outdated: totalOutdated,
    reviewed: totalReviewed,
    total: grandTotal,
    pctforReview: grandTotal ? totalReview / grandTotal : 0,
    pctOutdated: grandTotal ? totalOutdated / grandTotal : 0,
    pctReviewed: grandTotal ? totalReviewed / grandTotal : 0
  });

  return tableRows;
}

function getISOComplianceStatus() {
  var ss = SpreadsheetApp.openByUrl(appSettingURL);
  var sheet = ss.getSheetByName("ISO27001 Compliance");
 
  var data = sheet.getRange("A2:B3").getValues();

  iso_27001_compliance_data.push(data[0][1]);
  iso_27001_compliance_data.push(data[1][1]);
  
}