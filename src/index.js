// var xlsx = require("xlsx");
// var wb = xlsx.readFile("../inputfile/sample.xlsx", {});
// var ws = wb.Sheets["Pre program run"];
// var data = xlsx.utils.sheet_to_json(ws);
// console.log(data);
// //console.log(ws);
// var ws_data = [["category"]];
var cate = [
  "Category",
  "SHOPIFY",
  "SHOPIFY",
  "NOT_WORKING",
  "WOOCOMMERCE",
  "BIGCOMMERCE",
  "OTHERS",
  "MAGENTO",
];
// //var ws1 = xlsx.utils.aoa_to_sheet(ws_data);
// //let i = 0;
// // var newData = data.map(function (record) {
// //   record.category = cate[i];
// //   i++;
// // });
// // // console.log(newData);
// // wb.Sheets["Pre program run"] = ws1;
// console.log(ws.getColumn);
var Excel = require("exceljs");
var workbook = new Excel.Workbook();

workbook.xlsx.readFile("sample.xlsx").then(function () {
  var worksheet = workbook.getWorksheet(1);

  for (let i = 1; i <= 8; i++) {
    var row = worksheet.getRow(i);
    row.getCell(2).value = cate[i - 1];
    if (i === 1) {
      row.getCell(2).font = {
        size: 10,
        bold: true,
      };
    } else {
      row.getCell(2).font = {
        size: 10,
      };
    }
    row.commit();
  }

  return workbook.xlsx.writeFile("new.xlsx");
});
