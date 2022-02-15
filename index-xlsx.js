/* xlsx.js (C) 2013-present SheetJS -- https://sheetjs.com */
// const XLSX = require('xlsx');
const XLSX = require('sheetjs-style');

let results = {};
const processWb = (wb, fileName) => {
  console.log('wb: ', wb);
  wb.SheetNames.forEach(sheetName => {
    if (!results[sheetName]) {
      results[sheetName] = { ...wb, SheetNames: [], Sheets: {}};
      delete results[sheetName].Strings;
      delete results[sheetName].Workbook;
    }
    const res = results[sheetName];
    res.SheetNames.push(fileName);
    res.Sheets[fileName] = wb.Sheets[sheetName];
  });
};

const readFile = file => {
  const fileName = file.name.split('.')[0];
  const res = XLSX.readFile(file.path, {
    WTF: true,
    sheetStubs: true,
    cellDates: true,
    cellFormula: true,
    cellHTML: true,
    cellNF: true,
    cellStyles: true,
    cellText: true
  });
  processWb(res, fileName);
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
    XLSX.writeFile(results[key], `./out/${key}.xlsx`, {
      WTF: true,
      bookSST: true,
      bookVBA: true,
      cellDates: true,
      cellStyles: true
    });
  });
  results = {};
};

// add event listeners
const readIn = document.getElementById('readIn');
const exportBtn = document.getElementById('exportBtn');
const drop = document.getElementById('drop');

readIn.addEventListener('change', e => { readFiles(e.target.files); }, false);
exportBtn.addEventListener('click', exportXlsx, false);
drop.addEventListener('click', () => {
  readIn.click();
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
