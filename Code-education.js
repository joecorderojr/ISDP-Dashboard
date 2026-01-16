function groupAndCountDocStatusData() {
  //Logger.log(JSON.stringify(document_statuses));
  // Corrected data array where index [0] is the Unit and [1] is the Status
  const employeeDocuments = document_statuses;

  // Define all possible status types to ensure they appear in the output with 0 count if missing
  const allStatuses = ["Reviewed", "For Review", "Outdated"];

  const unitStatusCounts = employeeDocuments.reduce(
    (accumulator, currentItem) => {
      // Accessing values by assumed index:
      const unitKey = currentItem[0]; // e.g., "HR"
      const statusKey = currentItem[1]; // e.g., "For Review"

      // --- First Level Grouping (Group by Unit) ---
      if (!accumulator[unitKey]) {
        // If the Unit is new, initialize its object and set all status counters to 0.
        accumulator[unitKey] = {};
        allStatuses.forEach((status) => {
          accumulator[unitKey][status] = 0;
        });
      }

      // --- Second Level Aggregation (Count by Status) ---
      // Safely check if the status is one we track, then increment the counter.
      if (accumulator[unitKey].hasOwnProperty(statusKey)) {
        accumulator[unitKey][statusKey]++;
      } else {
        // Log a warning if an unknown status is found, but don't count it.
        Logger.log(
          `Warning: Skipping unknown status: ${statusKey} in ${unitKey}`
        );
      }

      return accumulator;
    },
    {}
  );

  // Output the final result
  Logger.log("--- Grouped and Counted Data ---");
  Logger.log(JSON.stringify(unitStatusCounts, null, 2));

  return unitStatusCounts;
}

/**
 * Transforms the grouped count data into the structured format
 * required for chart libraries (like Chart.js).
 */
function pivotCountsToChartDocStatusData() {
  // 1. INPUT: Use the grouped and counted data structure from the previous step.
  // Note: These values are hardcoded for demonstration; in a real script, this would
  // be the variable returned by the groupAndCount function.
  const groupedCounts = document_matrix_raw;

  // 2. Define the exact Statuses and their properties (colors, labels)
  const statusDefinitions = [
    { label: "Reviewed", color: "#50297d" },
    { label: "For Review", color: "#87a6ff" },
    { label: "Outdated", color: "#ff7f0e" },
  ];

  // --- Step A: Extract Labels (Units) ---
  // Get all unique unit names to form the 'labels' array
  const units = Object.keys(groupedCounts).sort();

  // --- Step B: Build Datasets ---
  const datasets = statusDefinitions.map((statusDef) => {
    // For each status (Reviewed, For Review, Outdated), create one dataset object

    // 1. Initialize an empty array for the count data points
    const dataPoints = [];

    // 2. Populate the dataPoints array by iterating over ALL Units (labels)
    units.forEach((unit) => {
      const unitData = groupedCounts[unit];

      // Get the count for the current unit and current status.
      // Use 0 if the unit/status combination doesn't exist in the data (for safety).
      const count = unitData ? unitData[statusDef.label] || 0 : 0;

      dataPoints.push(count);
    });

    // 3. Construct the final dataset object
    return {
      label: statusDef.label,
      data: dataPoints,
      backgroundColor: statusDef.color,
      borderRadius: 4,
    };
  });

  // --- Step C: Assemble Final Object ---
  const DocMaindata = {
    labels: units,
    datasets: datasets,
  };

  //Logger.log("--- Final Chart Data Structure ---");
  //Logger.log(JSON.stringify(DocMaindata, null, 2));

  return DocMaindata;
}

function pivotCountsToChartDocStatusDataPercentage(filterStatus, customColor) {
  // 1. INPUT: Use the grouped and counted data structure from the previous step.
  const groupedCounts = document_matrix_raw;

  // Define the exact Statuses and their properties (default colors, labels)
  const ALL_STATUS_DEFS = [
    { label: "Reviewed", defaultColor: defaultBGColors.green },
    { label: "For Review", defaultColor: defaultBGColors.yellow },
    { label: "Outdated", defaultColor: defaultBGColors.red }
  ];

  // --- Step A: Calculate Total Count per Unit ---
  const unitTotals = {};
  for (const unit in groupedCounts) {
    if (groupedCounts.hasOwnProperty(unit)) {
      const counts = groupedCounts[unit];
      // Sum all status counts for the current unit
      unitTotals[unit] =
        (counts.Reviewed || 0) +
        (counts["For Review"] || 0) +
        (counts.Outdated || 0);
    }
  }

  // --- Step B: Determine which statuses to process and Extract Labels (Units) ---
  const units = Object.keys(groupedCounts).sort();

  let statusesToProcess;

  // LOGIC CHANGE: If filterStatus is null/undefined OR is not a valid status, show all three.
  if (!filterStatus || !ALL_STATUS_DEFS.some((d) => d.label === filterStatus)) {
    // Default case: Show all three statuses with default colors
    statusesToProcess = ALL_STATUS_DEFS;
    filterStatus = "All (Default)";
  } else {
    // Specific status requested: Find the single status definition
    const singleDef = ALL_STATUS_DEFS.find((d) => d.label === filterStatus);
    statusesToProcess = singleDef ? [singleDef] : [];
  }

  // --- Step C: Build Datasets with Percentages ---
  const datasets = statusesToProcess.map((statusDef) => {
    const dataPoints = [];

    units.forEach((unit) => {
      const total = unitTotals[unit];
      // Safely access the count using the status label as the key
      const count = groupedCounts[unit][statusDef.label] || 0;
      let percentage = 0;

      if (total > 0) {
        // Calculate percentage: (Count / Total) * 100
        percentage = Math.round((count / total) * 100);
      }

      dataPoints.push(percentage);
    });

    // Apply custom color ONLY if a single status was specifically requested AND a custom color was provided.
    // If it defaulted to 'All', use defaultColor.
    const isSingleFiltered = statusesToProcess.length === 1;
    const finalColor =
      isSingleFiltered && customColor ? customColor : statusDef.defaultColor;

    // Construct the final dataset object
    return {
      label: statusDef.label,
      data: dataPoints,
      backgroundColor: finalColor,
      borderRadius: 4,
    };
  });

  // --- Step D: Assemble Final Object ---
  const DocMaindata = {
    labels: units,
    datasets: datasets,
  };

  //Logger.log("--- Final Chart Data Structure (Filtered by: %s) ---", filterStatus);
  //Logger.log(JSON.stringify(DocMaindata, null, 2));

  return DocMaindata;
}

/**
 * Iterates through grouped unit data, calculates the total count for all
 * statuses, and adds a percentage field for each status relative to the total.
 * For Matrix
 */
function calculateTotalsAndPercentages() {
  const groupedCounts = document_matrix_raw;

  // Define the statuses that should be included in the total calculation
  const statuses = ["Reviewed", "For Review", "Outdated"];

  // Use Object.keys() to loop through each unit (HR, Admin, etc.)
  Object.keys(groupedCounts).forEach((unitKey) => {
    const unitData = groupedCounts[unitKey];
    let total = 0;

    // 1. Calculate the Total
    statuses.forEach((status) => {
      // Ensure the count exists (defaults to 0 if not present)
      total += unitData[status] || 0;
    });

    // Add the Total property to the unit's data object
    unitData.Total = total;

    // 2. Calculate Percentages
    statuses.forEach((status) => {
      const count = unitData[status] || 0;
      let percentage = 0;

      if (total > 0) {
        // Calculate (Count / Total) * 100 and round to a whole number for cleanliness
        percentage = Math.round((count / total) * 100);
      }

      // Add the Percentage property to the status key
      unitData[status + "Percentage"] = percentage;
    });
  });

  //Logger.log("--- Data with Totals and Percentages ---");
  //Logger.log(JSON.stringify(groupedCounts, null, 2));

  return Object.entries(groupedCounts);
}

/**
 * Creates a Chart.js data object for a Doughnut/Pie chart,
 * showing the total document usage (raw count) per organizational unit.
 * The results are sorted by count in descending order.
 * This function now relies on the pre-calculated 'Total' key in the raw data.
 * * @returns {Object} Chart.js data structure (labels, data, and distinct colors).
 */
function getChartJsTopDocumentUsageMetrics() {
  const groupedCounts = globalThis.document_matrix_raw || {};

  const tempMetrics = [];

  // Define a comprehensive list of distinct colors for the slices
  const COLORS = [
    "#004d99", // 1. Deep Blue
    "#ff7f0e", // 2. Vibrant Orange
    "#2ecc71", // 3. Emerald Green
    "#85c1e9", // 4. Light Sky Blue
    "#5d6d7e", // 5. Slate Gray
    "#a569bd", // 6. Amethyst Purple
    "#f4d03f", // 7. Golden Yellow
    "#7f8c8d", // 8. Soft Gray
    "#e74c3c", // 9. Crimson Red
    "#1abc9c", // 10. Turquoise
    "#f1c40f", // 11. Sun Yellow
    "#3498db", // 12. Bright Blue
    "#9b59b6", // 13. Light Purple
    "#16a085", // 14. Dark Teal
    "#d35400", // 15. Deep Rust
  ];

  // 1. Extract Total Documents per Unit (relying only on 'Total' key)
  Object.keys(groupedCounts).forEach((unitKey) => {
    const unitData = groupedCounts[unitKey];
    // Check if the pre-calculated 'Total' property exists and is a valid number
    const totalCount = unitData.Total || 0;

    if (totalCount > 0) {
      tempMetrics.push({
        unit: unitKey,
        count: totalCount,
      });
    }
  });

  // 2. Sort by Count (Descending)
  tempMetrics.sort((a, b) => b.count - a.count);

  // 3. Extract Labels, Data, and Assign Colors
  const labels = [];
  const dataCounts = [];
  const backgroundColors = [];

  tempMetrics.forEach((metric, index) => {
    labels.push(metric.unit);
    dataCounts.push(metric.count);

    // Cycle through the predefined color array
    backgroundColors.push(COLORS[index % COLORS.length]);
  });

  // 4. Assemble Final Chart.js Object
  const TDUdata = {
    labels: labels,
    datasets: [
      {
        data: dataCounts,
        backgroundColor: backgroundColors,
        hoverOffset: 4,
      },
    ],
  };

  // Logger.log("--- Chart.js Total Document Usage Data (Sorted) ---");
  //Logger.log(JSON.stringify(TDUdata, null, 2));

  return TDUdata;
}

function getDataPrivacyTrainingData2026() {
  var response = getDataPrivacyTrainingData("2026");
  dataPrivacyTrainings2026 = response.result;
  dataPrivacyTrainingsTotal2026 = response.total;
  dataPrivacyTrainingsList2026 = response.objList;
}

function getDataPrivacyTrainingData2025() {
  var response = getDataPrivacyTrainingData("2025");
  dataPrivacyTrainings2025 = response.result;
  dataPrivacyTrainingsTotal2025 = response.total;
  dataPrivacyTrainingsList2025 = response.objList;
}

function getDataPrivacyTrainingData(year) {
  var ss = SpreadsheetApp.openByUrl(appSettingURL);
  var sheet = ss.getSheetByName("DataPrivacy"+year);

  var dataArray = sheet.getRange("A1:G1000").getValues();
  const data = filterEmptyRows(dataArray);

  const headers = data[0];
  const rows = data.slice(1);

  const col = headers.reduce((map, h, i) => ((map[h] = i), map), {});

  const units = {};
  var objList = [];

  rows.forEach((row) => {
    const unit = row[col["Business Unit"]];
    if (!unit) return;

    const status = row[col["Status"]];
    const score = parseFloat(row[col["Score"]]) || null;

    if (!units[unit]) {
      units[unit] = { employees: 0, inProgress: 0, completed: 0, scores: [] };
    }

    units[unit].employees++;

    if (status === "In Progress") units[unit].inProgress++;
    if (status === "Completed") {
      units[unit].completed++;
      if (score !== null) units[unit].scores.push(score);
    }
    else {
      units[unit].inProgress++;
    }

    objList.push({
      businessunit: row[col["Business Unit"]],
      name: row[col["Name"]],
      position: row[col["Title"]],
      email: row[col["Email"]],
      status: status,
      score: score,
    });
  });

  let totalEmployees = 0;
  let totalInProgress = 0;
  let totalCompleted = 0;
  let allScores = [];

  var result = Object.keys(units).map((unit) => {
    const u = units[unit];

    totalEmployees += u.employees;
    totalInProgress += u.inProgress;
    totalCompleted += u.completed;
    allScores = allScores.concat(u.scores);

    const percentCompleted = u.employees
      ? ((u.completed / u.employees) * 100).toFixed(0) + "%"
      : "0%";

    const avgQuiz = u.scores.length
      ? (
          (u.scores.reduce((a, b) => a + b) / u.scores.length / 10) *
          100
        ).toFixed(0) + "%"
      : "0%";

    return {
      unit,
      employeeCount: u.employees,
      inProgressCount: u.inProgress,
      completedCount: u.completed,
      pctCompleted: percentCompleted,
      pctAvgQuiz: avgQuiz,
    };
  });

  const totalPercentCompleted = totalEmployees
    ? ((totalCompleted / totalEmployees) * 100).toFixed(0) + "%"
    : "0%";

  const totalAvgQuiz = allScores.length
    ? (
        (allScores.reduce((a, b) => a + b) / allScores.length / 10) *
        100
      ).toFixed(0) + "%"
    : "0%";

  result = sortArrayByProperty(result, "unit");

  var total = {
    unit: "Total",
    employeeCount: totalEmployees,
    inProgressCount: totalInProgress,
    completedCount: totalCompleted,
    pctCompleted: totalPercentCompleted,
    pctAvgQuiz: totalAvgQuiz,
  };

  result.push(total);

  return {
    result: result,
    total: total,
    objList: objList
  } 
}

function getCyberSecurityTrainingData2026() {
  var reponse = getCyberSecurityTrainingData(2026);

  cyberSecurityTrainings2026 = reponse.result;
  cyberSecurityTrainingsTotal2026 = reponse.total;
  cyberSecurityTrainingsList2026 = reponse.objList;
}

function getCyberSecurityTrainingData2025() {
  var reponse = getCyberSecurityTrainingData(2025);

  cyberSecurityTrainings2025 = reponse.result;
  cyberSecurityTrainingsTotal2025 = reponse.total;
  cyberSecurityTrainingsList2025 = reponse.objList;
}

function getCyberSecurityTrainingData2024() {
  var reponse = getCyberSecurityTrainingData(2024);

  cyberSecurityTrainings2024 = reponse.result;
  cyberSecurityTrainingsTotal2024 = reponse.total;
  cyberSecurityTrainingsList2024 = reponse.objList;
}

function getCyberSecurityTrainingData(year) {
  var ss = SpreadsheetApp.openByUrl(appSettingURL);
  var sheet = ss.getSheetByName("CyberSecurity" + year);
  var dataArray = sheet.getRange("A1:G1000").getValues();
  const data = filterEmptyRows(dataArray);

  const headers = data[0];
  const rows = data.slice(1);

  const col = headers.reduce((map, h, i) => ((map[h] = i), map), {});

  const units = {};
  var objList = [];

  rows.forEach((row) => {
    const unit = row[col["Business Unit"]];

    if (!unit) return;

    const status = row[col["Status"]];
    const score = parseFloat(row[col["Score"]]) || null;

    if (!units[unit]) {
      units[unit] = {
        employees: 0,
        inProgress: 0,
        completed: 0,
        completedLate: 0,
        incompleteMissing: 0,
        scores: [],
      };
    }

    units[unit].employees++;

    if (status === "In Progress") units[unit].inProgress++;
    if (status === "Incomplete - Missing") units[unit].incompleteMissing++;
    if (status === "Completed - Late") units[unit].completedLate++;
    if (status === "Completed") {
      units[unit].completed++;
      if (score !== null) units[unit].scores.push(score);
    }
    else {
      units[unit].inProgress++;
    }

    objList.push({
      businessunit: row[col["Business Unit"]],
      name: row[col["Name"]],
      position: row[col["Title"]],
      email: row[col["Email"]],
      status: status,
      score: score,
    });
  });

  let totalEmployees = 0;
  let totalInProgress = 0;
  let totalCompleted = 0;
  let totalIncompleteMissing = 0;
  let totalCompleteLate = 0;
  let allScores = [];

  var result = Object.keys(units).map((unit) => {
    const u = units[unit];

    totalEmployees += u.employees;
    totalInProgress += u.inProgress;
    totalIncompleteMissing += u.incompleteMissing;
    totalCompleteLate += u.completedLate;
    totalCompleted += u.completed;
    allScores = allScores.concat(u.scores);

    const percentCompleted = u.employees
      ? ((u.completed / u.employees) * 100).toFixed(0) + "%"
      : "0%";

    const avgQuiz = u.scores.length
      ? (u.scores.reduce((a, b) => a + b) / u.scores.length).toFixed(0) + "%"
      : "0%";

    return {
      unit,
      employeeCount: u.employees,
      inProgressCount: u.inProgress,
      incompleteMissing: u.incompleteMissing,
      completedLate: u.completedLate,
      completedCount: u.completed,
      completedTotal: u.completed + u.completedLate,
      pctCompleted: percentCompleted,
      pctAvgQuiz: avgQuiz,
    };
  });

  const totalPercentCompleted = totalEmployees
    ? (((totalCompleted + totalCompleteLate) / totalEmployees) * 100).toFixed(
        0
      ) + "%"
    : "0%";

  const totalPercentIncompleteMissing = totalIncompleteMissing
    ? ((totalIncompleteMissing / totalEmployees) * 100).toFixed(0) + "%"
    : "0%";

  const totalAvgQuiz = allScores.length
    ? (allScores.reduce((a, b) => a + b) / allScores.length).toFixed(0) + "%"
    : "0%";

  result = sortArrayByProperty(result, "unit");

  var total = {
    unit: "Total",
    employeeCount: totalEmployees,
    inProgressCount: totalInProgress,
    incompleteMissing: totalIncompleteMissing,
    completedLate: totalCompleteLate,
    completedCount: totalCompleted,
    completedTotal: totalCompleteLate + totalCompleted,
    pctCompleted: totalPercentCompleted,
    pctAvgQuiz: totalAvgQuiz,
  };
  result.push(total);

  return { result: result, total: total, objList: objList };
}

function getPhishingResult() {
  var ss = SpreadsheetApp.openByUrl(appSettingURL);
  var sheet = ss.getSheetByName("Phishing Result");
  var dataArray = sheet.getRange("A1:E20").getValues();
  const data = filterEmptyRows(dataArray);

  const headers = data[0];
  const rows = data.slice(1);

  Logger.log(rows);

  const col = headers.reduce((map, h, i) => ((map[h] = i), map), {});

  const labels = [];
  const emailSent = [];
  //const reportedEmail = [];
  //const clickedLink = [];
  const submittedData = [];
  rows.forEach((row) => {
    labels.push(row[col["Period"]]);
    emailSent.push(row[col["Email Sent"]]);
    //reportedEmail.push(row[col["Reported Email"]]);
    //clickedLink.push(row[col["Clicked Link"]]);
    submittedData.push(row[col["Submitted Data"]]);
  });

  var result = {
    labels: labels,
    emailSent: emailSent,
    //reportedEmail : reportedEmail,
    //clickedLink : clickedLink,
    submittedData: submittedData,
  };

  phishingResult = result;
}
