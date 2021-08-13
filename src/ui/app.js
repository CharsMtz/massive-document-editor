const { remote } = require('electron')
const path = require('path')
const main = remote.require('./main.js')

const generateForm = document.getElementById('generateForm');
const dataFileInput = document.getElementById('fileDatos');
const templateFileInput = document.getElementById('filePlantilla');

generateForm.addEventListener('submit', (e) =>{
    const dataFilePath = dataFileInput.files[0].path;
    const templateFilePath = templateFileInput.files[0].path;
    const templateFileExtension= path.extname(templateFilePath);
    console.log(templateFileExtension);
    main.createDirectory().then(() => {
        main.generateDocs(dataFilePath, templateFilePath, templateFileExtension).then(()=>{
            main.showNotification();
        })
     })
})
