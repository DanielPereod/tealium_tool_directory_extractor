const fs = require('fs');
const UglifyJS = require("uglify-js");

const jsFilePath = './main.js';

fs.readFile(jsFilePath, 'utf8', (err, data) => {
    if (err) {
        console.error('Error al leer el archivo JavaScript:', err);
        return;
    }
    const compiledCode = UglifyJS.minify(data);

    const jsonContent = {
        id: "download_tealium_tree",
        title: "Get Tealium Server-side tree",
        description: "",
        icon_url: "https://tealium-tools.s3.amazonaws.com/tools/logo.png",
        scripts: {
            message: {
                title: "Download excel file",
                description: "",
                button_label: "Download",
                js: compiledCode.code 
            }
        }
    };

    const outputFilePath = './tealium_extractor.json';
    fs.writeFile(outputFilePath, JSON.stringify(jsonContent, null, 4), (err) => {
        if (err) {
            console.error('Error al escribir el archivo JSON:', err);
            return;
        }
        console.log('Archivo JSON creado exitosamente.');
    });
});
