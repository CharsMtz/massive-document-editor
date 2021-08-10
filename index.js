const xlsx = require('xlsx')
const path = require('path')
const Pizzip = require('pizzip');
const DocxTemplater = require('docxtemplater');
const fs = require('fs');

function generateDocs(file){
    const libro = xlsx.readFile(file);
    const hojasExcel = libro.SheetNames;
    const hoja = hojasExcel[0];
    const data = xlsx.utils.sheet_to_json(libro.Sheets[hoja]);

    for(let i=0; i<data.length;i++){
        var content = fs.readFileSync(path.resolve(__dirname, 'plantilla.docx'), 'binary');
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
        fs.writeFileSync(`Docs/${i+1}.docx`, buf);
    }
}

generateDocs('datos.xlsx');