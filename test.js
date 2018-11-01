var XlsxTemplate, data, fs, outputFile, sheetNumber, template, templateFile;

XlsxTemplate = require('./index.js');

fs = require('fs-extra');

data = fs.readJsonSync('./test/array.json');

templateFile = fs.readFileSync('./test/template.xlsx');

template = new XlsxTemplate(templateFile);

sheetNumber = 1;

template.substitute(sheetNumber, {
  rows: data
});

outputFile = template.generate();

fs.writeFileSync('./test/output.xlsx', outputFile, 'binary');
