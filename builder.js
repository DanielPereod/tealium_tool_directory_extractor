const fs = require('fs');
const UglifyJS = require("uglify-js");

const mainJsFilePath = './main.js';
const excelJsFilePath = 'exceljs.min.js';

Promise.all([
    fs.promises.readFile(mainJsFilePath, 'utf8'),
    fs.promises.readFile(excelJsFilePath, 'utf8')
])
.then(([mainJsData, excelJsData]) => {
    // Combine the contents of both files
    const combinedCode = excelJsData + '\n' + mainJsData;

    // Minify the combined code
    const compiledCode = UglifyJS.minify(combinedCode);

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
    const rawCodeFilePath = './tealium_extractor_raw.js';
    fs.promises.writeFile(outputFilePath, JSON.stringify(jsonContent, null, 4));
    return fs.promises.writeFile(rawCodeFilePath, combinedCode);

})
.then(() => {
    console.log('Archivo JSON creado exitosamente.');
})
.catch((err) => {
    console.error('Error:', err);
});
