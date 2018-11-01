###jshint globalstrict:true, devel:true ###

###eslint no-var:0 ###

###global require, module, Buffer ###

'use strict'
fs = require('fs-extra')
path = require('path')
zip = require('jszip')
etree = require('elementtree')
module.exports = do ->
  DOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
  CALC_CHAIN_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'
  SHARED_STRINGS_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
  HYPERLINK_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

  ###*
  # Create a new workbook. Either pass the raw data of a .xlsx file,
  # or call `loadTemplate()` later.
  ###

  Workbook = (data) ->
    self = this
    self.archive = null
    self.sharedStrings = []
    self.sharedStringsLookup = {}
    if data
      self.loadTemplate data
    return

  _get_simple = (obj, desc) ->
    if desc.indexOf('[') >= 0
      specification = desc.split(/[[[\]]/)
      property = specification[0]
      index = specification[1]
      return obj[property][index]
    obj[desc]

  ###*
  # Based on http://stackoverflow.com/questions/8051975
  # Mimic https://lodash.com/docs#get
  ###

  _get = (obj, desc, defaultValue) ->
    arr = desc.split('.')
    try
      while arr.length
        obj = _get_simple(obj, arr.shift())
    catch ex

      ### invalid chain ###

      obj = undefined
    if obj == undefined then defaultValue else obj

  ###*
  * Delete unused sheets if needed
  ###

  Workbook::deleteSheet = (sheetName) ->
    self = this
    sheet = self.loadSheet(sheetName)
    sh = self.workbook.find('sheets/sheet[@sheetId=\'' + sheet.id + '\']')
    self.workbook.find('sheets').remove sh
    rel = self.workbookRels.find('Relationship[@Id=\'' + sh.attrib['r:id'] + '\']')
    self.workbookRels.remove rel
    self._rebuild()
    self

  ###*
  * Clone sheets in current workbook template
  ###

  Workbook::copySheet = (sheetName, copyName) ->
    self = this
    sheet = self.loadSheet(sheetName)
    #filename, name , id, root
    newSheetIndex = (self.workbook.findall('sheets/sheet').length + 1).toString()
    fileName = 'worksheets' + '/' + 'sheet' + newSheetIndex + '.xml'
    arcName = self.prefix + '/' + fileName
    self.archive.file arcName, etree.tostring(sheet.root)
    self.archive.files[arcName].options.binary = true
    newSheet = etree.SubElement(self.workbook.find('sheets'), 'sheet')
    newSheet.attrib.name = copyName or 'Sheet' + newSheetIndex
    newSheet.attrib.sheetId = newSheetIndex
    newSheet.attrib['r:id'] = 'rId' + newSheetIndex
    newRel = etree.SubElement(self.workbookRels, 'Relationship')
    newRel.attrib.Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
    newRel.attrib.Target = fileName
    self._rebuild()
    #    TODO: work with "definedNames"
    #    var defn = etree.SubElement(self.workbook.find('definedNames'), 'definedName');
    #
    self

  ###*
  *  Partially rebuild after copy/delete sheets
  ###

  Workbook::_rebuild = ->
    #each <sheet> 'r:id' attribute in '\xl\workbook.xml'
    #must point to correct <Relationship> 'Id' in xl\_rels\workbook.xml.rels
    self = this
    order = [
      'worksheet'
      'theme'
      'styles'
      'sharedStrings'
    ]
    self.workbookRels.findall('*').sort((rel1, rel2) ->
      #using order
      index1 = order.indexOf(path.basename(rel1.attrib.Type))
      index2 = order.indexOf(path.basename(rel2.attrib.Type))
      if index1 + index2 == 0
        if rel1.attrib.Id and rel2.attrib.Id
          return rel1.attrib.Id.substring(3) - rel2.attrib.Id.substring(3)
        return rel1._id - (rel2._id)
      index1 - index2
    ).forEach (item, index) ->
      item.attrib.Id = 'rId' + index + 1
      return
    self.workbook.findall('sheets/sheet').forEach (item, index) ->
      item.attrib['r:id'] = 'rId' + index + 1
      item.attrib.sheetId = (index + 1).toString()
      return
    self.archive.file self.prefix + '/' + '_rels' + '/' + path.basename(self.workbookPath) + '.rels', etree.tostring(self.workbookRels)
    self.archive.file self.workbookPath, etree.tostring(self.workbook)
    self.sheets = self.loadSheets(self.prefix, self.workbook, self.workbookRels)
    return

  ###*
  # Load a .xlsx file from a byte array.
  ###

  Workbook::loadFile = (path)->
    data = fs.readFileSync(path)
    @loadTemplate(data)


  Workbook::writeFile = (path)->
    fs.writeFileSync path, @generate(),  'binary'

  Workbook::loadTemplate = (data) ->
    self = this
    if Buffer.isBuffer(data)
      data = data.toString('binary')
    self.archive = new zip(data,
      base64: false
      checkCRC32: true)
    # Load relationships
    rels = etree.parse(self.archive.file('_rels/.rels').asText()).getroot()
    workbookPath = rels.find('Relationship[@Type=\'' + DOCUMENT_RELATIONSHIP + '\']').attrib.Target
    self.workbookPath = workbookPath
    self.prefix = path.dirname(workbookPath)
    self.workbook = etree.parse(self.archive.file(workbookPath).asText()).getroot()
    self.workbookRels = etree.parse(self.archive.file(self.prefix + '/' + '_rels' + '/' + path.basename(workbookPath) + '.rels').asText()).getroot()
    self.sheets = self.loadSheets(self.prefix, self.workbook, self.workbookRels)
    self.calChainRel = self.workbookRels.find('Relationship[@Type=\'' + CALC_CHAIN_RELATIONSHIP + '\']')
    if self.calChainRel
      self.calcChainPath = self.prefix + '/' + self.calChainRel.attrib.Target
    self.sharedStringsPath = self.prefix + '/' + self.workbookRels.find('Relationship[@Type=\'' + SHARED_STRINGS_RELATIONSHIP + '\']').attrib.Target
    self.sharedStrings = []
    etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot().findall('si').forEach (si) ->
      t = text: ''
      si.findall('t').forEach (tmp) ->
        t.text += tmp.text
        return
      si.findall('r/t').forEach (tmp) ->
        t.text += tmp.text
        return
      self.sharedStrings.push t.text
      self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1
      return
    return

  ###*
  # Interpolate values for the sheet with the given number (1-based) or
  # name (if a string) using the given substitutions (an object).
  ###

  Workbook::substitute = (sheetName, substitutions) ->
    self = this
    sheet = self.loadSheet(sheetName)
    dimension = sheet.root.find('dimension')
    sheetData = sheet.root.find('sheetData')
    currentRow = null
    totalRowsInserted = 0
    totalColumnsInserted = 0
    namedTables = self.loadTables(sheet.root, sheet.filename)
    rows = []
    sheetData.findall('row').forEach (row) ->
      row.attrib.r = currentRow = self.getCurrentRow(row, totalRowsInserted)
      rows.push row
      cells = []
      cellsInserted = 0
      newTableRows = []
      row.findall('c').forEach (cell) ->
        appendCell = true
        cell.attrib.r = self.getCurrentCell(cell, currentRow, cellsInserted)
        # If c[@t="s"] (string column), look up /c/v@text as integer in
        # `this.sharedStrings`



        if cell.attrib.t == 's'
          # Look for a shared string that may contain placeholders
          cellValue = cell.find('v')
          stringIndex = parseInt(cellValue.text, 10)
          string = self.sharedStrings[stringIndex]
          if string == undefined
            return
          # Loop over placeholders
          self.extractPlaceholders(string).forEach (placeholder) ->
            # Only substitute things for which we have a substitution
            substitution = _get(substitutions, placeholder.name, '')
            newCellsInserted = 0
            if placeholder.full and placeholder.type == 'table' and substitution instanceof Array
              newCellsInserted = self.substituteTable(row, newTableRows, cells, cell, namedTables, substitution, placeholder.key)
              # don't double-insert cells
              # this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
              if newCellsInserted != 0 or substitution.length
                if substitution.length == 1
                  appendCell = true
                if substitution[0][placeholder.key] instanceof Array
                  appendCell = false
              # Did we insert new columns (array values)?
              if newCellsInserted != 0
                cellsInserted += newCellsInserted
                self.pushRight self.workbook, sheet.root, cell.attrib.r, newCellsInserted
            else if placeholder.full and placeholder.type == 'normal' and substitution instanceof Array
              appendCell = false
              # don't double-insert cells
              newCellsInserted = self.substituteArray(cells, cell, substitution)
              if newCellsInserted != 0
                cellsInserted += newCellsInserted
                self.pushRight self.workbook, sheet.root, cell.attrib.r, newCellsInserted
            else
              if placeholder.key
                substitution = _get(substitutions, placeholder.name + '.' + placeholder.key)
              string = self.substituteScalar(cell, string, placeholder, substitution)
            return
        # if we are inserting columns, we may not want to keep the original cell anymore
        if appendCell
          cells.push cell
        return
      # cells loop
      # We may have inserted columns, so re-build the children of the row
      self.replaceChildren row, cells
      # Update row spans attribute
      if cellsInserted != 0
        self.updateRowSpan row, cellsInserted
        if cellsInserted > totalColumnsInserted
          totalColumnsInserted = cellsInserted
      # Add newly inserted rows
      if newTableRows.length > 0
        newTableRows.forEach (row) ->
          rows.push row
          ++totalRowsInserted
          return
        self.pushDown self.workbook, sheet.root, namedTables, currentRow, newTableRows.length
      return
    # rows loop
    # We may have inserted rows, so re-build the children of the sheetData
    self.replaceChildren sheetData, rows
    # Update placeholders in table column headers
    self.substituteTableColumnHeaders namedTables, substitutions
    # Update placeholders in hyperlinks
    self.substituteHyperlinks sheet.filename, substitutions
    # Update <dimension /> if we added rows or columns
    if dimension
      if totalRowsInserted > 0 or totalColumnsInserted > 0
        dimensionRange = self.splitRange(dimension.attrib.ref)
        dimensionEndRef = self.splitRef(dimensionRange.end)
        dimensionEndRef.row += totalRowsInserted
        dimensionEndRef.col = self.numToChar(self.charToNum(dimensionEndRef.col) + totalColumnsInserted)
        dimensionRange.end = self.joinRef(dimensionEndRef)
        dimension.attrib.ref = self.joinRange(dimensionRange)
    #Here we are forcing the values in formulas to be recalculated
    # existing as well as just substituted
    sheetData.findall('row').forEach (row) ->
      row.findall('c').forEach (cell) ->
        formulas = cell.findall('f')
        if formulas and formulas.length > 0
          cell.findall('v').forEach (v) ->
            cell.remove v
            return
        return
      return
    # Write back the modified XML trees
    self.archive.file sheet.filename, etree.tostring(sheet.root)
    self.archive.file self.workbookPath, etree.tostring(self.workbook)
    # Remove calc chain - Excel will re-build, and we may have moved some formulae
    if self.calcChainPath and self.archive.file(self.calcChainPath)
      self.archive.remove self.calcChainPath
    self.writeSharedStrings()
    self.writeTables namedTables
    return

  ###*
  # Generate a new binary .xlsx file
  ###

  Workbook::generate = (options) ->
    self = this
    if !options
      options = base64: false
    self.archive.generate options

  # Helpers
  # Write back the new shared strings list

  Workbook::writeSharedStrings = ->
    self = this
    root = etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot()
    children = root.getchildren()
    root.delSlice 0, children.length
    self.sharedStrings.forEach (string) ->
      si = new (etree.Element)('si')
      t = new (etree.Element)('t')
      t.text = string
      si.append t
      root.append si
      return
    root.attrib.count = self.sharedStrings.length
    root.attrib.uniqueCount = self.sharedStrings.length
    self.archive.file self.sharedStringsPath, etree.tostring(root)
    return

  # Add a new shared string

  Workbook::addSharedString = (s) ->
    self = this
    idx = self.sharedStrings.length
    self.sharedStrings.push s
    self.sharedStringsLookup[s] = idx
    idx

  # Get the number of a shared string, adding a new one if necessary.

  Workbook::stringIndex = (s) ->
    self = this
    idx = self.sharedStringsLookup[s]
    if idx == undefined
      idx = self.addSharedString(s)
    idx

  # Replace a shared string with a new one at the same index. Return the
  # index.

  Workbook::replaceString = (oldString, newString) ->
    self = this
    idx = self.sharedStringsLookup[oldString]
    if idx == undefined
      idx = self.addSharedString(newString)
    else
      self.sharedStrings[idx] = newString
      delete self.sharedStringsLookup[oldString]
      self.sharedStringsLookup[newString] = idx
    idx

  # Get a list of sheet ids, names and filenames

  Workbook::loadSheets = (prefix, workbook, workbookRels) ->
    sheets = []
    workbook.findall('sheets/sheet').forEach (sheet) ->
      sheetId = sheet.attrib.sheetId
      relId = sheet.attrib['r:id']
      relationship = workbookRels.find('Relationship[@Id=\'' + relId + '\']')
      filename = prefix + '/' + relationship.attrib.Target
      sheets.push
        id: parseInt(sheetId, 10)
        name: sheet.attrib.name
        filename: filename
      return
    sheets

  # Get sheet a sheet, including filename and name

  Workbook::loadSheet = (sheet) ->
    self = this
    info = null
    i = 0
    while i < self.sheets.length
      if typeof sheet == 'number' and self.sheets[i].id == sheet or self.sheets[i].name == sheet
        info = self.sheets[i]
        break
      ++i
    if info == null and typeof sheet == 'number'
      #Get the sheet that corresponds to the 0 based index if the id does not work
      info = self.sheets[sheet - 1]
    if info == null
      throw new Error('Sheet ' + sheet + ' not found')
    {
      filename: info.filename
      name: info.name
      id: info.id
      root: etree.parse(self.archive.file(info.filename).asText()).getroot()
    }

  # Load tables for a given sheet

  Workbook::loadTables = (sheet, sheetFilename) ->
    self = this
    sheetDirectory = path.dirname(sheetFilename)
    sheetName = path.basename(sheetFilename)
    relsFilename = sheetDirectory + '/' + '_rels' + '/' + sheetName + '.rels'
    relsFile = self.archive.file(relsFilename)
    tables = []
    # [{filename: ..., root: ....}]
    if relsFile == null
      return tables
    rels = etree.parse(relsFile.asText()).getroot()
    sheet.findall('tableParts/tablePart').forEach (tablePart) ->
      relationshipId = tablePart.attrib['r:id']
      target = rels.find('Relationship[@Id=\'' + relationshipId + '\']').attrib.Target
      tableFilename = target.replace('..', self.prefix)
      tableTree = etree.parse(self.archive.file(tableFilename).asText())
      tables.push
        filename: tableFilename
        root: tableTree.getroot()
      return
    tables

  # Write back possibly-modified tables

  Workbook::writeTables = (tables) ->
    self = this
    tables.forEach (namedTable) ->
      self.archive.file namedTable.filename, etree.tostring(namedTable.root)
      return
    return

  #Perform substitution in hyperlinks

  Workbook::substituteHyperlinks = (sheetFilename, substitutions) ->
    self = this
    sheetDirectory = path.dirname(sheetFilename)
    sheetName = path.basename(sheetFilename)
    relsFilename = sheetDirectory + '/' + '_rels' + '/' + sheetName + '.rels'
    relsFile = self.archive.file(relsFilename)
    etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot()
    if relsFile == null
      return
    rels = etree.parse(relsFile.asText()).getroot()
    relationships = rels._children
    newRelationships = []
    relationships.forEach (relationship) ->
      newRelationships.push relationship
      if relationship.attrib.Type == HYPERLINK_RELATIONSHIP
        target = relationship.attrib.Target
        #Double-decode due to excel double encoding url placeholders
        target = decodeURI(decodeURI(target))
        self.extractPlaceholders(target).forEach (placeholder) ->
          substitution = substitutions[placeholder.name]
          if substitution == undefined
            return
          target = target.replace(placeholder.placeholder, self.stringify(substitution))
          relationship.attrib.Target = encodeURI(target)
          return
      return
    self.replaceChildren rels, newRelationships
    self.archive.file relsFilename, etree.tostring(rels)
    return

  # Perform substitution in table headers

  Workbook::substituteTableColumnHeaders = (tables, substitutions) ->
    self = this
    tables.forEach (table) ->
      `var tableRange`
      `var autoFilter`
      root = table.root
      columns = root.find('tableColumns')
      autoFilter = root.find('autoFilter')
      tableRange = self.splitRange(root.attrib.ref)
      idx = 0
      inserted = 0
      newColumns = []
      columns.findall('tableColumn').forEach (col) ->
        ++idx
        col.attrib.id = Number(idx).toString()
        newColumns.push col
        name = col.attrib.name
        self.extractPlaceholders(name).forEach (placeholder) ->
          substitution = substitutions[placeholder.name]
          if substitution == undefined
            return
          # Array -> new columns
          if placeholder.full and placeholder.type == 'normal' and substitution instanceof Array
            substitution.forEach (element, i) ->
              newCol = col
              if i > 0
                newCol = self.cloneElement(newCol)
                newCol.attrib.id = Number(++idx).toString()
                newColumns.push newCol
                ++inserted
                tableRange.end = self.nextCol(tableRange.end)
              newCol.attrib.name = self.stringify(element)
              return
            # Normal placeholder
          else
            name = name.replace(placeholder.placeholder, self.stringify(substitution))
            col.attrib.name = name
          return
        return
      self.replaceChildren columns, newColumns
      # Update range if we inserted columns
      if inserted > 0
        columns.attrib.count = Number(idx).toString()
        root.attrib.ref = self.joinRange(tableRange)
        if autoFilter != null
          # XXX: This is a simplification that may stomp on some configurations
          autoFilter.attrib.ref = self.joinRange(tableRange)
      #update ranges for totalsRowCount
      tableRoot = table.root
      tableRange = self.splitRange(tableRoot.attrib.ref)
      tableStart = self.splitRef(tableRange.start)
      tableEnd = self.splitRef(tableRange.end)
      if tableRoot.attrib.totalsRowCount
        autoFilter = tableRoot.find('autoFilter')
        if autoFilter != null
          autoFilter.attrib.ref = self.joinRange(
            start: self.joinRef(tableStart)
            end: self.joinRef(tableEnd))
        ++tableEnd.row
        tableRoot.attrib.ref = self.joinRange(
          start: self.joinRef(tableStart)
          end: self.joinRef(tableEnd))
      return
    return

  # Return a list of tokens that may exist in the string.
  # Keys are: `placeholder` (the full placeholder, including the `${}`
  # delineators), `name` (the name part of the token), `key` (the object key
  # for `table` tokens), `full` (boolean indicating whether this placeholder
  # is the entirety of the string) and `type` (one of `table` or `cell`)

  Workbook::extractPlaceholders = (string) ->
    # Yes, that's right. It's a bunch of brackets and question marks and stuff.
    re = /\{{(?:(.+?):)?(.+?)(?:\.(.+?))?}}/g
    match = null
    matches = []
    while (match = re.exec(string)) != null
      matches.push
        placeholder: match[0]
        type: match[1] or 'normal'
        name: match[2]
        key: match[3]
        full: match[0].length == string.length
    matches

  # Split a reference into an object with keys `row` and `col` and,
  # optionally, `table`, `rowAbsolute` and `colAbsolute`.

  Workbook::splitRef = (ref) ->
    match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/)
    {
      table: match and match[1] or null
      colAbsolute: Boolean(match and match[2])
      col: match and match[3]
      rowAbsolute: Boolean(match and match[4])
      row: parseInt(match and match[5], 10)
    }

  # Join an object with keys `row` and `col` into a single reference string

  Workbook::joinRef = (ref) ->
    (if ref.table then ref.table + '!' else '') + (if ref.colAbsolute then '$' else '') + ref.col.toUpperCase() + (if ref.rowAbsolute then '$' else '') + Number(ref.row).toString()

  # Get the next column's cell reference given a reference like "B2".

  Workbook::nextCol = (ref) ->
    self = this
    ref = ref.toUpperCase()
    ref.replace /[A-Z]+/, (match) ->
      self.numToChar self.charToNum(match) + 1

  # Get the next row's cell reference given a reference like "B2".

  Workbook::nextRow = (ref) ->
    ref = ref.toUpperCase()
    ref.replace /[0-9]+/, (match) ->
      (parseInt(match, 10) + 1).toString()

  # Turn a reference like "AA" into a number like 27

  Workbook::charToNum = (str) ->
    num = 0
    idx = str.length - 1
    iteration = 0
    while idx >= 0
      thisChar = str.charCodeAt(idx) - 64
      multiplier = 26 ** iteration
      num += multiplier * thisChar
      --idx
      ++iteration
    num

  # Turn a number like 27 into a reference like "AA"

  Workbook::numToChar = (num) ->
    str = ''
    i = 0
    while num > 0
      remainder = num % 26
      charCode = remainder + 64
      num = (num - remainder) / 26
      # Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
      if remainder == 0
        # 26 -> Z
        charCode = 90
        --num
      str = String.fromCharCode(charCode) + str
      ++i
    str

  # Is ref a range?

  Workbook::isRange = (ref) ->
    ref.indexOf(':') != -1

  # Is ref inside the table defined by startRef and endRef?

  Workbook::isWithin = (ref, startRef, endRef) ->
    self = this
    start = self.splitRef(startRef)
    end = self.splitRef(endRef)
    target = self.splitRef(ref)
    start.col = self.charToNum(start.col)
    end.col = self.charToNum(end.col)
    target.col = self.charToNum(target.col)
    start.row <= target.row and target.row <= end.row and start.col <= target.col and target.col <= end.col

  # Turn a value of any type into a string

  Workbook::stringify = (value) ->
    if value instanceof Date
      #In Excel date is a number of days since 01/01/1900
      #           timestamp in ms    to days      + number of days from 1900 to 1970
      return Number(value.getTime() / (1000 * 60 * 60 * 24) + 25569)
    else if typeof value == 'number' or typeof value == 'boolean'
      return Number(value).toString()
    else if typeof value == 'string'
      return String(value).toString()
    ''

  # Insert a substitution value into a cell (c tag)

  Workbook::insertCellValue = (cell, substitution) ->
    self = this
    cellValue = cell.find('v')
    stringified = self.stringify(substitution)
    if typeof substitution == 'string' and substitution[0] == '='
      #substitution, started with '=' is a formula substitution
      formula = new (etree.Element)('f')
      formula.text = substitution.substr(1)
      cell.insert 1, formula
      delete cell.attrib.t
      #cellValue will be deleted later
      return formula.text
    if typeof substitution == 'number' or substitution instanceof Date
      delete cell.attrib.t
      cellValue.text = stringified
    else if typeof substitution == 'boolean'
      cell.attrib.t = 'b'
      cellValue.text = stringified
    else
      cell.attrib.t = 's'
      cellValue.text = Number(self.stringIndex(stringified)).toString()
    stringified


  # Sheets with formulas that have tokens will produce #VALUE! errors because the tokens are text
  # This will strip the value attribute from the cell to ensure that the sheet recalcs on load



  # Perform substitution of a single value

  Workbook::substituteScalar = (cell, string, placeholder, substitution) ->
    self = this
    if placeholder.full
      self.insertCellValue cell, substitution
    else
      newString = string.replace(placeholder.placeholder, self.stringify(substitution))
      cell.attrib.t = 's'
      self.insertCellValue cell, newString

  # Perform a columns substitution from an array

  Workbook::substituteArray = (cells, cell, substitution) ->
    self = this
    newCellsInserted = -1
    currentCell = cell.attrib.r
    # add a cell for each element in the list
    substitution.forEach (element) ->
      ++newCellsInserted
      if newCellsInserted > 0
        currentCell = self.nextCol(currentCell)
      newCell = self.cloneElement(cell)
      self.insertCellValue newCell, element
      newCell.attrib.r = currentCell
      cells.push newCell
      return
    newCellsInserted

  # Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
  # Returns total number of new cells inserted on the original row.

  Workbook::substituteTable = (row, newTableRows, cells, cell, namedTables, substitution, key) ->
    self = this
    newCellsInserted = 0
    # on the original row
    # if no elements, blank the cell, but don't delete it
    if substitution.length == 0
      delete cell.attrib.t
      self.replaceChildren cell, []
    else
      parentTables = namedTables.filter((namedTable) ->
        range = self.splitRange(namedTable.root.attrib.ref)
        self.isWithin cell.attrib.r, range.start, range.end
      )
      substitution.forEach (element, idx) ->
        newRow = undefined
        newCell = undefined
        newCellsInsertedOnNewRow = 0
        newCells = []
        value = _get(element, key, '')
        if idx == 0
          # insert in the row where the placeholders are
          if value instanceof Array
            newCellsInserted = self.substituteArray(cells, cell, value)
          else
            self.insertCellValue cell, value
        else
          # insert new rows (or reuse rows just inserted)
          # Do we have an existing row to use? If not, create one.
          if idx - 1 < newTableRows.length
            newRow = newTableRows[idx - 1]
          else
            newRow = self.cloneElement(row, false)
            newRow.attrib.r = self.getCurrentRow(row, newTableRows.length + 1)
            newTableRows.push newRow
          # Create a new cell
          newCell = self.cloneElement(cell)
          newCell.attrib.r = self.joinRef(
            row: newRow.attrib.r
            col: self.splitRef(newCell.attrib.r).col)
          if value instanceof Array
            newCellsInsertedOnNewRow = self.substituteArray(newCells, newCell, value)
            # Add each of the new cells created by substituteArray()
            newCells.forEach (newCell) ->
              newRow.append newCell
              return
            self.updateRowSpan newRow, newCellsInsertedOnNewRow
          else
            self.insertCellValue newCell, value
            # Add the cell that previously held the placeholder
            newRow.append newCell
          # expand named table range if necessary
          parentTables.forEach (namedTable) ->
            tableRoot = namedTable.root
            autoFilter = tableRoot.find('autoFilter')
            range = self.splitRange(tableRoot.attrib.ref)
            if !self.isWithin(newCell.attrib.r, range.start, range.end)
              range.end = self.nextRow(range.end)
              tableRoot.attrib.ref = self.joinRange(range)
              if autoFilter != null
                # XXX: This is a simplification that may stomp on some configurations
                autoFilter.attrib.ref = tableRoot.attrib.ref
            return
        return
    newCellsInserted

  # Clone an element. If `deep` is true, recursively clone children

  Workbook::cloneElement = (element, deep) ->
    self = this
    newElement = etree.Element(element.tag, element.attrib)
    newElement.text = element.text
    newElement.tail = element.tail
    if deep != false
      element.getchildren().forEach (child) ->
        newElement.append self.cloneElement(child, deep)
        return
    newElement

  # Replace all children of `parent` with the nodes in the list `children`

  Workbook::replaceChildren = (parent, children) ->
    parent.delSlice 0, parent.len()
    children.forEach (child) ->
      parent.append child
      return
    return

  # Calculate the current row based on a source row and a number of new rows
  # that have been inserted above

  Workbook::getCurrentRow = (row, rowsInserted) ->
    parseInt(row.attrib.r, 10) + rowsInserted

  # Calculate the current cell based on asource cell, the current row index,
  # and a number of new cells that have been inserted so far

  Workbook::getCurrentCell = (cell, currentRow, cellsInserted) ->
    self = this
    colRef = self.splitRef(cell.attrib.r).col
    colNum = self.charToNum(colRef)
    self.joinRef
      row: currentRow
      col: self.numToChar(colNum + cellsInserted)

  # Adjust the row `spans` attribute by `cellsInserted`

  Workbook::updateRowSpan = (row, cellsInserted) ->
    if cellsInserted != 0 and row.attrib.spans
      rowSpan = row.attrib.spans.split(':').map((f) ->
        parseInt f, 10
      )
      rowSpan[1] += cellsInserted
      row.attrib.spans = rowSpan.join(':')
    return

  # Split a range like "A1:B1" into {start: "A1", end: "B1"}

  Workbook::splitRange = (range) ->
    split = range.split(':')
    {
      start: split[0]
      end: split[1]
    }

  # Join into a a range like "A1:B1" an object like {start: "A1", end: "B1"}

  Workbook::joinRange = (range) ->
    range.start + ':' + range.end

  # Look for any merged cell or named range definitions to the right of
  # `currentCell` and push right by `numCols`.

  Workbook::pushRight = (workbook, sheet, currentCell, numCols) ->
    self = this
    cellRef = self.splitRef(currentCell)
    currentRow = cellRef.row
    currentCol = self.charToNum(cellRef.col)
    # Update merged cells on the same row, at a higher column
    sheet.findall('mergeCells/mergeCell').forEach (mergeCell) ->
      mergeRange = self.splitRange(mergeCell.attrib.ref)
      mergeStart = self.splitRef(mergeRange.start)
      mergeStartCol = self.charToNum(mergeStart.col)
      mergeEnd = self.splitRef(mergeRange.end)
      mergeEndCol = self.charToNum(mergeEnd.col)
      if mergeStart.row == currentRow and currentCol < mergeStartCol
        mergeStart.col = self.numToChar(mergeStartCol + numCols)
        mergeEnd.col = self.numToChar(mergeEndCol + numCols)
        mergeCell.attrib.ref = self.joinRange(
          start: self.joinRef(mergeStart)
          end: self.joinRef(mergeEnd))
      return
    # Named cells/ranges
    workbook.findall('definedNames/definedName').forEach (name) ->
      ref = name.text
      if self.isRange(ref)
        namedRange = self.splitRange(ref)
        namedStart = self.splitRef(namedRange.start)
        namedStartCol = self.charToNum(namedStart.col)
        namedEnd = self.splitRef(namedRange.end)
        namedEndCol = self.charToNum(namedEnd.col)
        if namedStart.row == currentRow and currentCol < namedStartCol
          namedStart.col = self.numToChar(namedStartCol + numCols)
          namedEnd.col = self.numToChar(namedEndCol + numCols)
          name.text = self.joinRange(
            start: self.joinRef(namedStart)
            end: self.joinRef(namedEnd))
      else
        namedRef = self.splitRef(ref)
        namedCol = self.charToNum(namedRef.col)
        if namedRef.row == currentRow and currentCol < namedCol
          namedRef.col = self.numToChar(namedCol + numCols)
          name.text = self.joinRef(namedRef)
      return
    return

  # Look for any merged cell, named table or named range definitions below
  # `currentRow` and push down by `numRows` (used when rows are inserted).

  Workbook::pushDown = (workbook, sheet, tables, currentRow, numRows) ->
    self = this
    mergeCells = sheet.find('mergeCells')
    # Update merged cells below this row
    sheet.findall('mergeCells/mergeCell').forEach (mergeCell) ->
      mergeRange = self.splitRange(mergeCell.attrib.ref)
      mergeStart = self.splitRef(mergeRange.start)
      mergeEnd = self.splitRef(mergeRange.end)
      if mergeStart.row > currentRow
        mergeStart.row += numRows
        mergeEnd.row += numRows
        mergeCell.attrib.ref = self.joinRange(
          start: self.joinRef(mergeStart)
          end: self.joinRef(mergeEnd))
      #add new merge cell
      if mergeStart.row == currentRow
        i = 1
        while i <= numRows
          newMergeCell = self.cloneElement(mergeCell)
          mergeStart.row += 1
          mergeEnd.row += 1
          newMergeCell.attrib.ref = self.joinRange(
            start: self.joinRef(mergeStart)
            end: self.joinRef(mergeEnd))
          mergeCells.attrib.count += 1
          mergeCells._children.push newMergeCell
          i++
      return
    # Update named tables below this row
    tables.forEach (table) ->
      tableRoot = table.root
      tableRange = self.splitRange(tableRoot.attrib.ref)
      tableStart = self.splitRef(tableRange.start)
      tableEnd = self.splitRef(tableRange.end)
      if tableStart.row > currentRow
        tableStart.row += numRows
        tableEnd.row += numRows
        tableRoot.attrib.ref = self.joinRange(
          start: self.joinRef(tableStart)
          end: self.joinRef(tableEnd))
        autoFilter = tableRoot.find('autoFilter')
        if autoFilter != null
          # XXX: This is a simplification that may stomp on some configurations
          autoFilter.attrib.ref = tableRoot.attrib.ref
      return
    # Named cells/ranges
    workbook.findall('definedNames/definedName').forEach (name) ->
      ref = name.text
      if self.isRange(ref)
        namedRange = self.splitRange(ref)
        namedStart = self.splitRef(namedRange.start)
        namedEnd = self.splitRef(namedRange.end)
        if namedStart
          if namedStart.row > currentRow
            namedStart.row += numRows
            namedEnd.row += numRows
            name.text = self.joinRange(
              start: self.joinRef(namedStart)
              end: self.joinRef(namedEnd))
      else
        namedRef = self.splitRef(ref)
        if namedRef.row > currentRow
          namedRef.row += numRows
          name.text = self.joinRef(namedRef)
      return
    return

  Workbook
