const { remote } = require('electron')
const main = remote.require('./main.js')
const fs = require('fs');

const generateForm = document.getElementById('generateForm');
const dataFileInput = document.getElementById('fileDatos');
const templateFileInput = document.getElementById('filePlantilla');




generateForm.addEventListener('submit', () =>{
    const dataFilePath = dataFileInput.files[0].path;
    const templateFilePath = templateFileInput.files[0].path;

    var copyDocs = main.copyDocs(dataFilePath, templateFilePath).then(() => {
        var createDirectory = main.createDirectory().then(() => {
            main.generateDocs();
        })
    })
})
