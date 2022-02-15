/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* from the electron quick-start */
const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const XLSX = require('xlsx');

let win = null;

function createWindow() {
  if (win) return;
  win = new BrowserWindow({
    width: 1200, height: 900,
    webPreferences: {
      worldSafeExecuteJavaScript: true, // required for Electron 12+
      contextIsolation: false, // required for Electron 12+
      nodeIntegration: true,
      enableRemoteModule: true
    }
  });
  win.loadFile('index.html');
  win.webContents.openDevTools();
  win.on('closed', function () { win = null; });
}
if (app.setAboutPanelOptions) app.setAboutPanelOptions({ applicationName: 'sheetjs-electron', applicationVersion: 'XLSX ' + XLSX.version, copyright: '(C) 2017-present SheetJS LLC' });
app.on('open-file', function () { console.log(arguments); });
app.on('ready', createWindow);
app.on('activate', createWindow);
app.on('window-all-closed', function () { if (process.platform !== 'darwin') app.quit(); });

ipcMain.on('open-file-dialog', event => {
  dialog.showOpenDialog({ properties: ['openDirectory'] }).then(res => {
    if (!res.canceled) {
      event.sender.send('selected-directory', res.filePaths[0]);
    }
  });
});
