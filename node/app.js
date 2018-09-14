var Excel = require('exceljs');

var workbook = new Excel.Workbook();
workbook.xlsx.readFile('./test.xlsx')
.then(function () {
    var worksheet = workbook.getWorksheet(1)
    var c = worksheet.getCell('A46')
    console.log(c.value)
    worksheet.eachRow((row,rowNumber) => {
        let pRow = rowNumber
        // row.eachCell((cell) => {
        //     pRow+=cell.value+' '
        // });
        console.log(pRow)
    });
});
