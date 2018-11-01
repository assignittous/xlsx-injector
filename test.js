var XlsxTemplate, data, fs, sheetNumber, template;

XlsxTemplate = require('./index.js');

fs = require('fs-extra');

data = fs.readJsonSync('./test/array.json');

//templateFile = fs.readFileSync()
template = new XlsxTemplate;

template.loadFile('./test/template.xlsx');

sheetNumber = 1;

template.sheets.forEach(function(sheet) {
  return template.substitute(sheet.id, {
    rows: data
  });
});

template.writeFile('./test/output.xlsx');
