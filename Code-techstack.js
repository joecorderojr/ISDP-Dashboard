const getDynamicColors = (data, type) => {
    return data.map(value => {
        if (type === 'contract') {
            if (value === 100) return 'rgba(75, 192, 192, 0.8)'; // Green (Renewed)
            if (value === 50) return 'rgba(255, 205, 86, 0.8)';  // Yellow (Expiring)
            return 'rgba(255, 99, 132, 0.8)';                  // Red (Terminated)
        }
        // Status and Update use Green for 100, Red for 0
        return value === 100 ? 'rgba(75, 192, 192, 0.8)' : 'rgba(255, 99, 132, 0.8)';
    });
};

const getBarColors = (value) => {
      if (value === 'Active' || value === 'Updated' || value === 'Renewed') return 'barGreen'; 
      if (value === 'Inactive' || value === 'Not Updated' || value === 'Terminated') return 'barRed'; 
      if (value === 'Expiring') return 'barYellow'; 
};



function getTPSData() {
  try {
    // 1. Open the spreadsheet and get the data
    var ss = SpreadsheetApp.openByUrl(appSettingURL);
    const sheet = ss.getSheetByName('Tools'); // Ensure your sheet is named 'Sheet1'
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null; // Return null if there's no data below header

    // Get data from row 2, columns 1 and 2
    const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    // 2. Initialize the structure
    const labels = [];
    const statusData = [];
    const updateData = [];
    const contractData = [];

    // 3. Populate arrays based on Status text
    values.forEach(row => {
      const toolName = row[0];

      if(toolName)
      {
        labels.push(toolName);

        const status = String(row[1]).trim(); // Clean whitespace
        if (status === "Active") {
          statusData.push(100);  
        } else {
          statusData.push(0);
        }

        const update = String(row[2]).trim(); // Clean whitespace
        if (update === "Updated") {
          updateData.push(100);  
        } else {
          updateData.push(0);
        }

        const contract = String(row[3]).trim(); // Clean whitespace
        if (contract === "Renewed") {
          contractData.push(100);  
        } else if (contract === "Expiring") {
          contractData.push(50);  
        } else {
          contractData.push(0);
        }

      }

    });

    // 4. Construct the final object
    const TPSdata = {
      labels: labels,
      datasets: [
        {
            label: 'Status',
            data: statusData.map(() => 100), // Always 100% height
            backgroundColor: getDynamicColors(statusData, 'status'),
            borderWidth: 1
        },
        {
            label: 'Update',
            data: updateData.map(() => 100), // Always 100% height
            backgroundColor: getDynamicColors(updateData, 'update'),
            borderWidth: 1
        },
        {
            label: 'Contract',
            data: contractData.map(() => 100), // Always 100% height
            backgroundColor: getDynamicColors(contractData, 'contract'),
            borderWidth: 1
        }
    ]
    };

    Logger.log(TPSdata);

    tools_data = TPSdata;
    tools_matrix = values;

  } catch (e) {
    Logger.log('Error fetching data: ' + e.toString());
    return null;
  }
}

function getVASummaryData() {
  try {
    // 1. Open the spreadsheet and get the data
    var ss = SpreadsheetApp.openByUrl(appSettingURL);
    const sheet = ss.getSheetByName('VA Summary'); // Ensure your sheet is named 'Sheet1'
    
    // Get data from row 2, columns 1 and 2
    const values = sheet.getRange("B2:E2").getValues();

    const data = [values[0][0], values[0][1], values[0][2], values[0][3]];

    vaSummary_data = JSON.stringify(data, null, 2);

  } catch (e) {
    Logger.log('Error fetching data: ' + e.toString());
    return null;
  }
}


function getSOC2TYPE2Data() {
  try {
    // 1. Open the spreadsheet and get the data
    var ss = SpreadsheetApp.openByUrl(appSettingURL);
    const sheet = ss.getSheetByName('SOC2 TYPE2'); // Ensure your sheet is named 'Sheet1'
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null; // Return null if there's no data below header

    // Get data from row 2, columns 1 and 2
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    // 2. Initialize the structure
    const labels = [];
    const statusData = [];

    // 3. Populate arrays based on Status text
    values.forEach(row => {
      const labelName = row[0];
      if(labelName)
      {
        labels.push(labelName);
        const status = String(row[3]).trim(); // Clean whitespace
        statusData.push(status);   
      }

    });

    // 4. Construct the final object
    const TPSdata = {
      labels: labels,
      datasets: [
        {
            label: 'Status',
            data: statusData.map(() => 100), // Always 100% height
            backgroundColor: getDynamicColors(statusData, 'status'),
            borderWidth: 1
        },
        {
            label: 'Update',
            data: updateData.map(() => 100), // Always 100% height
            backgroundColor: getDynamicColors(updateData, 'update'),
            borderWidth: 1
        },
        {
            label: 'Contract',
            data: contractData.map(() => 100), // Always 100% height
            backgroundColor: getDynamicColors(contractData, 'contract'),
            borderWidth: 1
        }
    ]
    };

    Logger.log(TPSdata);

    tools_data = TPSdata;
    tools_matrix = values;

  } catch (e) {
    Logger.log('Error fetching data: ' + e.toString());
    return null;
  }
}
