const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const directoryPath = path.join(__dirname, "excels");

async function processExcelFiles() {
  const valueCounts = {};
  // read file in folder
  const files = fs.readdirSync(directoryPath);
  for (const file of files) {
    const filePath = path.join(directoryPath, file);
    if (path.extname(filePath) === ".xlsx") {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);

      const worksheet = workbook.getWorksheet(1);
      if (worksheet) {
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          // start row 2
          if (rowNumber > 1) {
            const cellValue = row.getCell(1).value;
            if (cellValue) {
              if (!valueCounts[cellValue]) {
                valueCounts[cellValue] = 0;
              }
              valueCounts[cellValue]++;
            }
          }
        });
      }
    }
  }

  const sortedValueCounts = Object.entries(valueCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([value, count]) => `${value}: ${count}`)
    .join("\n");

  const outputFilePath = path.join(__dirname, "output.txt");
  fs.writeFileSync(outputFilePath, sortedValueCounts);
}

processExcelFiles();
