
function uploadAndSummarize(base64Data, fileName, targetFolderId) {
  try {
    var period = fileName;
    fileName = "VA Summary " + fileName;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      MimeType.MICROSOFT_EXCEL,
      fileName
    );

    const resource = {
      title: fileName.replace(/\.(xlsx|xls)$/i, ""),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: targetFolderId }]
    };

    const convertedFile = Drive.Files.insert(resource, blob, { convert: true });

    summarizeRisk(convertedFile.id, period);

    return {
      fileUrl: convertedFile.alternateLink,
      name: convertedFile.title
    };

  } catch (err) {
    throw new Error(err.toString());
  }
}

function summarizeRisk(fileId, period) {
  const inputSheet = SpreadsheetApp.openById(fileId).getSheets()[0];
  const inputData = inputSheet.getDataRange().getValues();

  const dataRows = inputData.slice(1);

  const riskLevels = ["Critical", "High", "Medium", "Low"];
  const summary = {};
  const totals = { "Critical": 0, "High": 0, "Medium": 0, "Low": 0 };

  dataRows.forEach(row => {
    const risk = row[0];
    const host = row[1];

    if (!summary[host]) {
      summary[host] = { "Critical": 0, "High": 0, "Medium": 0, "Low": 0 };
    }

    if (riskLevels.includes(risk)) {
      summary[host][risk]++;
      totals[risk]++;
    }
  });

  const output = [["IP", ...riskLevels]];

  output.push([
    "",
    totals.Critical,
    totals.High,
    totals.Medium,
    totals.Low
  ]);

  Object.keys(summary).sort().forEach(host => {
    const row = [host];
    riskLevels.forEach(r => row.push(summary[host][r]));
    output.push(row);
  });

  const outputSpreadsheet = SpreadsheetApp.openByUrl(appSettingURL);
  let outputSheet = outputSpreadsheet.getSheetByName('VA Summary');

  if (!outputSheet) {
    outputSheet = outputSpreadsheet.insertSheet('VA Summary');
  } else {
    outputSheet.getRange("A2:E1000").clear();
  }

  const range = outputSheet.getRange(1, 1, output.length, output[0].length);
  range.setValues(output);
  outputSheet.getRange("H2").setValue(period);

  const lastCol = output[0].length;

  outputSheet
    .getRange(1, 1, 1, lastCol)
    .setBackground('#d5a6bd')
    .setFontWeight('bold');

  outputSheet
    .getRange(2, 2, 1, lastCol - 1)
    .setBackground('#cfe2f3')
    .setFontWeight('bold');
}
