/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("filter-table").onclick = filterTable;
    document.getElementById("sort-table").onclick = sortTable;
    document.getElementById("create-chart").onclick = createChart;
    document.getElementById("freeze-header").onclick = freezeHeader;
    document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("new-data").onclick = newData;
    document.getElementById("slicer").onclick = addSlicer;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function createTable() {
  Excel.run(function (context) {
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values =
      [["Date", "Merchant", "Category", "Amount"]];
    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['_ € * #.##0_ ;#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function filterTable() {
  Excel.run(function (context) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
      var categoryFilter = expensesTable.columns.getItem('Category').filter;
      categoryFilter.applyValuesFilter(['Education', 'Groceries']);
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
       }
  });
}

function sortTable() {
  Excel.run(function (context) {
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 0,            // Merchant column
            ascending: true,
        }
    ];
    expensesTable.sort.apply(sortFields);
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function createChart() {
  Excel.run(function (context) {
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in &euro;';
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function freezeHeader() {
  Excel.run(function (context) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.freezePanes.freezeRows(1);
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

var dialog = null;
function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    {height: 45, width: 55},
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
);
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

function newData() {
  Excel.run(function (context) {
    context.workbook.worksheets.getItemOrNullObject("Data").delete();
    const dataSheet = context.workbook.worksheets.add("Data");
    context.workbook.worksheets.getItemOrNullObject("Pivot").delete();
    const pivotSheet = context.workbook.worksheets.add("Pivot");

    const data = [
      ["Farm", "Type", "Classification", "Crates Sold at Farm", "Crates Sold Wholesale"],
      ["A Farms", "Lime", "Organic", 300, 2000],
      ["A Farms", "Lemon", "Organic", 250, 1800],
      ["A Farms", "Orange", "Organic", 200, 2200],
      ["B Farms", "Lime", "Conventional", 80, 1000],
      ["B Farms", "Lemon", "Conventional", 75, 1230],
      ["B Farms", "Orange", "Conventional", 25, 800],
      ["B Farms", "Orange", "Organic", 20, 500],
      ["B Farms", "Lemon", "Organic", 10, 770],
      ["B Farms", "Kiwi", "Conventional", 30, 300],
      ["B Farms", "Lime", "Organic", 50, 400],
      ["C Farms", "Apple", "Organic", 275, 220],
      ["C Farms", "Kiwi", "Organic", 200, 120],
      ["D Farms", "Apple", "Conventional", 100, 3000],
      ["D Farms", "Apple", "Organic", 80, 2800],
      ["E Farms", "Lime", "Conventional", 160, 2700],
      ["E Farms", "Orange", "Conventional", 180, 2000],
      ["E Farms", "Apple", "Conventional", 245, 2200],
      ["E Farms", "Kiwi", "Conventional", 200, 1500],
      ["F Farms", "Kiwi", "Organic", 100, 150],
      ["F Farms", "Lemon", "Conventional", 150, 270]
    ];

    const rangeToAnalyze = dataSheet.getRange("A1:E21");
    rangeToAnalyze.values = data;
    rangeToAnalyze.format.autofitColumns();

    pivotSheet.activate();
    const rangeToPlacePivot = pivotSheet.getRange("A2");
    pivotSheet.pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));
    pivotTable.layout.layoutType = "Tabular";

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function addSlicer() {
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}
