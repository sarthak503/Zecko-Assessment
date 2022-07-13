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

  return workbook.xlsx.writeFile("sample.xlsx");
});
