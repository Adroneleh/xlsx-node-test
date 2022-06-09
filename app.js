const http = require("http");
const fs = require("fs");
const ExcelJS = require("exceljs");

const hostname = "127.0.0.1";
const port = 3000;

const server = http.createServer((req, res) => {
  res.statusCode = 200;
  res.setHeader("Content-Type", "text/plain");
  res.end("Hello World");
});

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});

const func = () => {
  console.time("write");

  const options = {
    worksheets: "emit",
  };
  const path = "path/to/file";

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(path, options);
  workbookReader.read();

  workbookReader.on("worksheet", (worksheet) => {
    worksheet.on("row", (row) => {
      fs.appendFileSync("result.json", JSON.stringify(row.values));
    });
  });
  workbookReader.on("end", () => {
    console.timeEnd("write");
    server.close();
  });
  workbookReader.on("error", (err) => {
    console.log(err);
  });
};

func();
