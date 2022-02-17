const fs = require('fs');
const { join } = require('path');
const { execSync } = require('child_process');
const Excel = require('exceljs');
const cloneDeep = require('lodash.clonedeep');

/** @type {Record<string, Excel.Workbook>} */
let results = {};
let selFolderPath = process.argv;
console.log('selFolderPath: ', selFolderPath);

const processWb = (/** @type {Excel.Workbook} */ wb, fileName) => {
  wb.eachSheet(sheet => {
    const sheetName = sheet.name;
    if (!results[sheetName]) {
      results[sheetName] = new Excel.Workbook();
    }
    const res = results[sheetName];
    const copySheet = res.addWorksheet(fileName);
    copySheet.model = cloneDeep(sheet.model);
    copySheet.name = fileName;
  });
};

const readFile = (filePath, fileName) => {
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile(filePath).then(wb => {
    processWb(wb, fileName);
  });
};

const exportXlsx = () => {
  const outPath = join(selFolderPath, '_out');
  if (!fs.existsSync(outPath)) fs.mkdirSync(outPath);
  Object.keys(results).forEach(key => {
    results[key].xlsx.writeFile(join(outPath, `${key}.xlsx`), {
      useSharedStrings: true,
      useStyles: true
    });
  });
  results = {};
  execSync('start ' + outPath);
};

const main = () => {
  try {
    execSync('convert-xls-xlsx.vbs ' + selFolderPath);
    const files = fs.readdirSync(selFolderPath);
    for (const filePath of files) {
      if (filePath.endsWith('.xlsx')) {
        readFile(join(selFolderPath, filePath), filePath.split('.')[0]);
      }
    }
    exportXlsx();
  } catch (error) {
    console.log('error: ', error);
  }
};

main();
