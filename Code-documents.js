function loadDocumentData()
{
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("Updated Masterlist");

  var lastrow = ws.getRange("A1").getDataRegion().getLastRow();
  var list = ws.getRange(2,1,lastrow,14).getValues();

  return list;
}

function initializedDocumentVariables()
{
  var types = [];
  var units = [];
  var status = [];
  let total_reviewed = 0;
  let total_for_review = 0;
  let total_overdue = 0;
  document_list.forEach(function(list, index) {
    if(!types.includes(list[3]))
    {
      types.push(list[3]);
    }

    if(!units.includes(list[0]))
    {
      units.push(list[0]);
    }

    switch (list[12]) {
      case "Overdue":
          total_overdue++;
          break;
      case "For Review":
          total_for_review++;
          break;
      case "Reviewed":
          total_reviewed++;
          break;
    }

    status.push([list[0],list[12]]);
 
  });

  data_doc_total.push(total_overdue);
  data_doc_total.push(total_reviewed);
  data_doc_total.push(total_for_review);

  document_types = types;
  unit_list = units;
  document_statuses = status;
  document_matrix_raw = groupAndCountDocStatusData();
  document_matrix = calculateTotalsAndPercentages();
 

 // Logger.log(JSON.stringify(document_matrix));

  const datasetObject = {
      label: "Reviewed", // First column is the label
      data: [0, 50, 25, 100, 13, 100, 91], // Remaining columns are the data points
      backgroundColor: '#000000', // Set color based on some logic
      borderRadius: 4
    };

  document_list.forEach(function(list, index) {
    if(!types.includes(list[3]))
    {
      types.push(list[3]);
    } 
  });

}