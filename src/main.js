const { BrowserWindow,app } = require('electron')
const xlsx = require('xlsx')
const path = require('path')
const Pizzip = require('pizzip');
const DocxTemplater = require('docxtemplater');
const fs = require('fs');


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
    const libro = xlsx.readFile('src/Docs/input/datos.xlsx');
    const hojasExcel = libro.SheetNames;
    const hoja = hojasExcel[0];
    const data = xlsx.utils.sheet_to_json(libro.Sheets[hoja]);
    return data;
}

function generateDocs(){
    const data = readExcel();
    for(let i=0; i<data.length;i++){
        var content = fs.readFileSync(path.resolve(__dirname, 'Docs/input/plantilla.docx'), 'binary');
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
        fs.writeFileSync(`src/Docs/output/${i+1}.docx`, buf);
    }
}




module.exports = {
    createWindow,
    generateDocs
}