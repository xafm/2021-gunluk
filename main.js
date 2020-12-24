const electron = require('electron');
const {app, BrowserWindow} = electron;
const path = require('path');
const url = require('url');

let win;

function createWindow() {
  const {width, height} = electron.screen.getPrimaryDisplay().workAreaSize;

  win = new BrowserWindow({
    width,
    height,
    frame: false,
    webPreferences: {
      nodeIntegration: true,
      enableRemoteModule: true,
    },
  });

  win.loadURL(
    url.format({
      pathname: path.join(__dirname, 'view/index.html'),
      protocol: 'file:',
      slashes: true,
    }),
  );

  win.on('closed', () => {
    win = null;
  });

  win.webContents.on('new-window', function(e, url) {
    e.preventDefault();
    require('electron').shell.openExternal(url);
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (win === null) {
    createWindow();
  }
});
