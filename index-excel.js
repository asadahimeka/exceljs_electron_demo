const { ipcRenderer } = require('electron');
const Excel = require('exceljs');
const _ = require('lodash');

const readIn = document.getElementById('readIn');
const exportBtn = document.getElementById('exportBtn');
const drop = document.getElementById('drop');

/** @type {Record<string, Excel.Workbook>} */
let results = {};

ipcRenderer.on('selected-directory', (_event, path) => {
  console.log('path: ', path);
});

const processWb = (/** @type {Excel.Workbook} */ wb, fileName) => {
  console.log('wb: ', wb);
  wb.eachSheet(sheet => {
    const sheetName = sheet.name;
    if (!results[sheetName]) {
      results[sheetName] = new Excel.Workbook();
    }
    const res = results[sheetName];
    const copySheet = res.addWorksheet(fileName);
    copySheet.model = _.cloneDeep(sheet.model);
    copySheet.name = fileName;
  });
};

const readFile = file => {
  const fileName = file.name.split('.')[0];
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile(file.path).then(wb => {
    processWb(wb, fileName);
  });
};

const readFiles = files => {
  console.log('files: ', files);
  for (let i = 0; i < files.length; i++) {
    readFile(files[i]);
  }
};

const exportXlsx = () => {
  console.log('results: ', results);
  Object.keys(results).forEach(key => {
    results[key].xlsx.writeFile(`./out/${key}.xlsx`, {
      useSharedStrings: true,
      useStyles: true
    });
  });
  results = {};
  readIn.value = null;
};

readIn.addEventListener('change', e => { readFiles(e.target.files); }, false);
exportBtn.addEventListener('click', exportXlsx, false);
drop.addEventListener('click', () => {
  // readIn.click();
  ipcRenderer.send('open-file-dialog');
});
drop.addEventListener('dragenter', e => {
  e.stopPropagation();
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
}, false);
drop.addEventListener('dragover', e => {
  e.stopPropagation();
  e.preventDefault();
  e.dataTransfer.dropEffect = 'copy';
}, false);
drop.addEventListener('drop', e => {
  e.stopPropagation();
  e.preventDefault();
  readFiles(e.dataTransfer.files);
}, false);
