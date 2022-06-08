const http = require("http");
// const XLSX = require("xlsx");
const fs = require("fs");
// const readXlsxFile = require("read-excel-file/node");
// const { convert } = require("xlsx-converter");

const hostname = "127.0.0.1";
const port = 3000;

const server = http.createServer((req, res) => {
  res.statusCode = 200;
  res.setHeader("Content-Type", "text/plain");
  res.end("Hello World");
});

// var ExcelCSV = require('excelcsv');

// var header  = ['id', 'email']
//   , fileIn  = '/home/adroneleh/Downloads/1.xlsx'
//   , fileOut = '/result.csv';

// // fileOut is optional.
// var parser = new ExcelCSV( fileIn, fileOut );
// var csv = parser
//             // optional.
//             .header( header )
//             // optional.
//             .row(function(worksheet, row) {
//               // Transform data here or return false to skip the row.
//               return row;
//             })
//             .init();

// var Excel = require('exceljs');
// var workbook = new Excel.Workbook(); 
// workbook.xlsx.readFile('/home/adroneleh/Downloads/300k+.xlsx')
//     .then(function() {
//         var worksheet = workbook.getWorksheet('Sheet1');
//         worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
//           console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
//         });
//     });

// convert("/home/adroneleh/Downloads/1.xlsx", { to: "csv" }).then((result) => {
//   // do something with the result
//   console.log(result);
// });

// const xlsx = require('node-xlsx');
// Or var xlsx = require('node-xlsx').default;

// Parse a buffer
// Parse a file
// const workSheetsFromFile = xlsx.parse(`/home/adroneleh/Downloads/300k+.xlsx`);
// console.log(workSheetsFromFile)
// console.time("reading");

const Excel = require('exceljs');
const func = async () => {
  const workbookReader = new Excel.stream.xlsx.WorkbookReader('fileName');
  for await (const worksheetReader of workbookReader) {
    for await (const row of worksheetReader) {
      // fs.appendFileSync("result.json", JSON.stringify(row));
      console.log(row)
    }
  }
}

func()

// const readable = fs.createReadStream("/home/adroneleh/Downloads/300k+.xlsx");

// const writeToFile = (result) => {
//   fs.appendFileSync("result.json", result);
// };

// async function process_RS(stream, callback) {
//   const workbook = new Excel.Workbook();
//   const worksheet = await workbook.xlsx.read(stream);

//   console.log(worksheet.worksheets[0]._rows)
//   // worksheet.eachSheet((ws, id) => {
//   //   console.log(ws)
//   // });
// }

// process_RS(readable, writeToFile);

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
