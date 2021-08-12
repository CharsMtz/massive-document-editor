const { remote } = require('electron')
const main = remote.require('./main.js')
const fs = require('fs');

const generateForm = document.getElementById('generateForm');
const dataFileInput = document.getElementById('fileDatos');
const templateFileInput = document.getElementById('filePlantilla');




generateForm.addEventListener('submit', (e) =>{
    e.preventDefault();
    const dataFilePath = dataFileInput.files[0].path;
    const templateFilePath = dataFileInput.files[0].path;
    fs.copyFile(templateFilePath, 'src/Docs/input/plantilla.docx', (err) => {
        if (err) throw err;
        console.log('doc was copied to destination.txt');
    });
    fs.copyFile(dataFilePath, 'src/Docs/input/datos.xlsx', (err) => {
        if (err) throw err;
        console.log('excel was copied to destination.txt');
    });
    main.generateDocs();
})
