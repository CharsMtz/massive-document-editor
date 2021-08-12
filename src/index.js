const { app } = require('electron')
const { createWindow } = require('./main')



require('electron-reload')(__dirname);

app.whenReady().then(createWindow);



//filefn.generateDocs('src/Docs/input/datos.xlsx');