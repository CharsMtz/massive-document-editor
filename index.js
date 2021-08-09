const fs = require('fs');
const path = require('path')
const Pizzip = require('pizzip');
const DocxTemplater = require('docxtemplater');

var data = fs.readFileSync(path.resolve(__dirname, 'invitados.csv')).toString();
var datosArray = data.split("\r\n");
console.log(datosArray);

for (let i = 0; i < datosArray.length; i++) {
    

    var campos = datosArray[i].split(",");

    if(campos[0]){
        var content = fs.readFileSync(path.resolve(__dirname, 'plantilla.docx'), 'binary');
        const zip = new Pizzip(content);
        const doc = new DocxTemplater(zip);
        doc.setData({
            id: campos[0],
            nombre: campos[1],
            puesto: campos[2],
            cuarto: campos[3]
        });
        try {
            doc.render();
        } catch (e) {
            throw e;
        }
        var buf = doc.getZip().generate({ type: 'nodebuffer' });
        fs.writeFileSync(`Docs/${campos[1]}.docx`, buf);
    }   
}