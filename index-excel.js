const fs = require('fs');
const { join } = require('path');
const { execSync } = require('child_process');
const { ipcRenderer } = require('electron');
const Excel = require('exceljs');
const cloneDeep = require('lodash.clonedeep');

/** @type {Record<string, Excel.Workbook>} */
let results = {};
let selFolderPath = '';

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
  try {
    execSync('start ' + outPath);
  } catch (err) {
    console.log('err: ', err);
  }
};

const selFolderBtn = document.getElementById('selfolder');
const exportBtn = document.getElementById('exportBtn');

ipcRenderer.on('selected-directory', (_event, path) => {
  try {
    selFolderPath = path;
    if (process.env.npm_lifecycle_event) {
      execSync('convert-xls-xlsx.vbs ' + path);
    } else {
      execSync(join(process.resourcesPath, 'app', 'convert-xls-xlsx.vbs') + ' ' + path);
    }
    const files = fs.readdirSync(path);
    for (const filePath of files) {
      if (filePath.endsWith('.xlsx')) {
        readFile(join(path, filePath), filePath.split('.')[0]);
      }
    }
  } catch (error) {
    console.log('error: ', error);
    alert('运行出错：' + error);
  }
});

exportBtn.addEventListener('click', exportXlsx, false);
selFolderBtn.addEventListener('click', () => {
  ipcRenderer.send('open-file-dialog');
});
