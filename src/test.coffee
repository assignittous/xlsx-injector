XlsxTemplate = require('./index.js')
fs = require('fs-extra')

data = fs.readJsonSync('./test/array.json')

template = new XlsxTemplate

template.loadFile './test/template.xlsx'

template.sheets.forEach (sheet)->
  template.substitute sheet.id,
    rows: data

template.writeFile './test/output.xlsx'
