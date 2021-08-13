const { BrowserWindow,app } = require('electron')
const xlsx = require('xlsx')
const Pizzip = require('pizzip');
const DocxTemplater = require('docxtemplater');
const fs = require('fs');
const { windowsStore } = require('process');

const homeDir = require('os').homedir();
const desktopDir = `${homeDir}/Desktop/Documentos_Generados`;

const { Notification } = require('electron')
const NOTIFICATION_TITLE = 'Archivos generados'
const NOTIFICATION_BODY = 'Creación de archivos exitosa. Loas archivos están dentro de la carpeta Documentos_Generados en el escritorio.'

function showNotification () {
  new Notification({ title: NOTIFICATION_TITLE, body: NOTIFICATION_BODY }).show()
}

let window

function createWindow(){
    window = new BrowserWindow({ 
        width:500,
        height: 350,
        show:false,
        webPreferences:{
            nodeIntegration: true,
            contextIsolation: false,
            enableRemoteModule:true
        }
    });
    window.removeMenu();
    window.setResizable(false);
    window.loadFile('src/ui/index.html');
    window.webContents.on('did-finish-load', function() {
        window.show();
    });
    window.on("closed", () => {
        app.quit()
    });
}

// Massive Document generator functions
function readExcel(filePath){
    const libro = xlsx.readFile(filePath);
    const hojasExcel = libro.SheetNames;
    const hoja = hojasExcel[0];
    const data = xlsx.utils.sheet_to_json(libro.Sheets[hoja]);
    return data;
}

async function generateDocs(dataFilePath, templateFilePath, templateFileExtension){
    const data = readExcel(dataFilePath);
    for(let i=0; i<data.length;i++){
        var content = fs.readFileSync(templateFilePath, 'binary');
        const zip = new Pizzip(content);
        const doc = new DocxTemplater(zip);
        doc.setData(
            data[i]
        );
        try {
            doc.render();
        } catch (e) {
            throw e;
        }
        var buf = doc.getZip().generate({ type: 'nodebuffer' });
        fs.writeFileSync(`${desktopDir}/${i+1}${templateFileExtension}`, buf);
    }
}

async function createDirectory(){
    console.log(desktopDir);
    fs.mkdir(desktopDir, { recursive: true }, function(err) {
        if (err) {
          console.log(err)
        } else {
          console.log("New directory successfully created.")
        }
      })
}

module.exports = {
    createWindow,
    generateDocs,
    createDirectory,
    showNotification
}