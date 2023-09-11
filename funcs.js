
function getInputValues() {
  // Find input fields, store them in a constant
  const inputFields = document.querySelectorAll("#myForm input[type='text']");
  // Create empty inputValues array
  const inputValues = [];
  // Loop through input fields
  for (let i = 0; i < inputFields.length; i++) {
    // Push values of each input field into an array
    inputValues.push(inputFields[i].value);
  }
  // Log array in the console
  console.log(inputValues);
  exportToCsv()
}

await Excel.run(async (context) => {
  let sheets = context.workbook.worksheets;
  sheets.load("items/name");

  await context.sync();
  
  if (sheets.items.length > 1) {
      console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
  } else {
      console.log(`There is one worksheet in the workbook:`);
  }

  sheets.items.forEach(function (sheet) {
      console.log(sheet.name);
  });
});

await Excel.run(async (context) => {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load("name");

  await context.sync();
  console.log(`The active worksheet is "${sheet.name}"`);
});

await Excel.run(async (context) => {
  let sheet = context.workbook.worksheets.getItem("Sample");
  sheet.activate();
  sheet.load("name");

  await context.sync();
  console.log(`The active worksheet is "${sheet.name}"`);
});

await Excel.run(async (context) => {
  let sheet = context.workbook.worksheets.getItem("Sample");
  sheet.visibility = Excel.SheetVisibility.hidden;
  sheet.load("name");

  await context.sync();
  console.log(`Worksheet with name "${sheet.name}" is hidden`);
});

/*
exportToCsv = function() {
  var CsvString = "";
  Results.forEach(function(RowItem, RowIndex) {
    RowItem.forEach(function(ColItem, ColIndex) {
      CsvString += ColItem + ',';
    });
    CsvString += "\r\n";
  });
  CsvString = "data:application/csv," + encodeURIComponent(CsvString);
  var x = document.createElement("A");
  x.setAttribute("href", CsvString );
  x.setAttribute("download", "data.csv");
  document.body.appendChild(x);
  x.click();
} */
