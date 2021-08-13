const { BrowserWindow,app } = require('electron')
const xlsx = require('xlsx')
const path = require('path')
const Pizzip = require('pizzip');
const DocxTemplater = require('docxtemplater');
const fs = require('fs');

const homeDir = require('os').homedir();
const desktopDir = `${homeDir}/Desktop/Documentos_Generados`;

let window

function createWindow(){
    window = new BrowserWindow({ 
        width:500,
        height: 480,
        webPreferences:{
            nodeIntegration: true,
            contextIsolation: false,
            enableRemoteModule:true
        }
    });
    window.loadFile('src/ui/index.html');
    window.show();
    window.on("closed", () => {
        app.quit()
    });
}

// Document generator functions
function readExcel(){
    const libro = xlsx.readFile('src/Docs/datos.xlsx');
    const hojasExcel = libro.SheetNames;
    const hoja = hojasExcel[0];
    const data = xlsx.utils.sheet_to_json(libro.Sheets[hoja]);
    return data;
}

async function generateDocs(){
    const data = readExcel();
    for(let i=0; i<data.length;i++){
        var content = fs.readFileSync(path.resolve(__dirname, 'Docs/plantilla.docx'), 'binary');
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
        fs.writeFileSync(`${desktopDir}/${i+1}.docx`, buf);
    }
}

async function copyDocs(dataFilePath, templateFilePath){
    fs.copyFile(templateFilePath, 'src/Docs/plantilla.docx', (err) => {
        if (err) throw err;
        console.log('template was copied');
    });
    fs.copyFile(dataFilePath, 'src/Docs/datos.xlsx', (err) => {
        if (err) throw err;
        console.log('excel was copied');
    });
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
    copyDocs,
    createDirectory
}