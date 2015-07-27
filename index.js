
/*jshint globalstrict:true, devel:true */

/*global require, module, exports, process, __dirname, Buffer */
'use strict';
var etree, fs, path, zip;

fs = require('fs');

path = require('path');

zip = require('node-zip');

etree = require('elementtree');

require('sugar');

module.exports = (function() {
  var DOCUMENT_RELATIONSHIP, SHARED_STRINGS_RELATIONSHIP, Workbook;
  DOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  SHARED_STRINGS_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';

  /**
   * Create a new workbook. Either pass the raw data of a .xlsx file,
   * or call `loadTemplate()` later.
   */
  Workbook = function(input_path) {
    var self;
    self = this;
    self.archive = null;
    self.sharedStrings = [];
    self.sharedStringsLookup = {};
    self.input_path = null;
    self.output_path = null;
    if (input_path) {
      self.loadFile(input_path);
    }
  };

  /**
   * Load a .xlsx file from a byte array.
   */
  Workbook.prototype.loadFile = function(path) {
    var data;
    data = fs.readFileSync(path);
    return this.loadTemplate(data);
  };
  Workbook.prototype.writeFile = function(path) {
    return fs.writeFileSync(path, this.generate(), 'binary');
  };
  Workbook.prototype.loadTemplate = function(data) {
    var rels, self, workbookPath;
    self = this;
    if (Buffer.isBuffer(data)) {
      data = data.toString('binary');
    }
    self.archive = new zip(data, {
      base64: false,
      checkCRC32: true
    });
    rels = etree.parse(self.archive.file('_rels/.rels').asText()).getroot();
    workbookPath = rels.find('Relationship[@Type=\'' + DOCUMENT_RELATIONSHIP + '\']').attrib.Target;
    self.workbookPath = workbookPath;
    self.prefix = path.dirname(workbookPath);
    self.workbook = etree.parse(self.archive.file(workbookPath).asText()).getroot();
    self.workbookRels = etree.parse(self.archive.file(self.prefix + '/_rels/' + path.basename(workbookPath) + '.rels').asText()).getroot();
    self.sheets = self.loadSheets(self.prefix, self.workbook, self.workbookRels);
    self.sharedStringsPath = self.prefix + '/' + self.workbookRels.find('Relationship[@Type=\'' + SHARED_STRINGS_RELATIONSHIP + '\']').attrib.Target;
    self.sharedStrings = [];
    etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot().findall('si/t').forEach(function(t) {
      self.sharedStrings.push(t.text);
      self.sharedStringsLookup[t.text] = self.sharedStrings.length - 1;
    });
  };
  Workbook.prototype.purgeBadValues = function(sheetName) {};

  /**
   * Interpolate values for the sheet with the given number (1-based) or
   * name (if a string) using the given substitutions (an object).
   */
  Workbook.prototype.substitute = function(sheetName, substitutions) {
    var currentRow, namedTables, rows, self, sheet, sheetData, totalRowsInserted;
    self = this;
    sheet = self.loadSheet(sheetName);
    sheetData = sheet.root.find('sheetData');
    currentRow = null;
    totalRowsInserted = 0;
    namedTables = self.loadTables(sheet.root, sheet.filename);
    rows = [];
    sheetData.findall('row').forEach(function(row) {
      var cells, cellsInserted, newTableRows;
      row.attrib.r = currentRow = self.getCurrentRow(row, totalRowsInserted);
      rows.push(row);
      cells = [];
      cellsInserted = 0;
      newTableRows = [];
      row.findall('c').forEach(function(cell) {
        var appendCell, cellValue, string, stringIndex;
        appendCell = true;
        cell.attrib.r = self.getCurrentCell(cell, currentRow, cellsInserted);
        if ((cell.attrib.t === 'e') || (cell.attrib.t === 'str')) {
          cellValue = cell.find('v');
          self.stripValue(cell);
        }
        if (cell.attrib.t === 's') {
          cellValue = cell.find('v');
          stringIndex = parseInt(cellValue.text, 10);
          string = self.sharedStrings[stringIndex];
          if (string === void 0) {
            return;
          }
          self.extractPlaceholders(string).forEach(function(placeholder) {
            var newCellsInserted, substitution;
            substitution = substitutions[placeholder.name];
            newCellsInserted = 0;
            if (substitution === void 0) {
              return;
            }
            if (placeholder.full && placeholder.type === 'table' && substitution instanceof Array) {
              newCellsInserted = self.substituteTable(row, newTableRows, cells, cell, namedTables, substitution, placeholder.key);
              if (newCellsInserted !== 0) {
                appendCell = false;
                cellsInserted += newCellsInserted;
                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
              }
            } else if (placeholder.full && placeholder.type === 'normal' && substitution instanceof Array) {
              appendCell = false;
              newCellsInserted = self.substituteArray(cells, cell, substitution);
              if (newCellsInserted !== 0) {
                cellsInserted += newCellsInserted;
                self.pushRight(self.workbook, sheet.root, cell.attrib.r, newCellsInserted);
              }
            } else {
              string = self.substituteScalar(cell, string, placeholder, substitution);
            }
          });
        }
        if (appendCell) {
          cells.push(cell);
        }
      });
      self.replaceChildren(row, cells);
      if (cellsInserted !== 0) {
        self.updateRowSpan(row, cellsInserted);
      }
      if (newTableRows.length > 0) {
        newTableRows.forEach(function(row) {
          rows.push(row);
          ++totalRowsInserted;
        });
        self.pushDown(self.workbook, sheet.root, namedTables, currentRow, newTableRows.length);
      }
    });
    self.replaceChildren(sheetData, rows);
    self.substituteTableColumnHeaders(namedTables, substitutions);
    self.archive.file(sheet.filename, etree.tostring(sheet.root));
    self.archive.file(self.workbookPath, etree.tostring(self.workbook));
    self.writeSharedStrings();
    self.writeTables(namedTables);
  };

  /**
   * Generate a new binary .xlsx file
   */
  Workbook.prototype.generate = function() {
    var self;
    self = this;
    return self.archive.generate({
      base64: false
    });
  };
  Workbook.prototype.writeSharedStrings = function() {
    var children, root, self;
    self = this;
    root = etree.parse(self.archive.file(self.sharedStringsPath).asText()).getroot();
    children = root.getchildren();
    root.delSlice(0, children.length);
    self.sharedStrings.forEach(function(string) {
      var si, t;
      si = new etree.Element('si');
      t = new etree.Element('t');
      t.text = string;
      si.append(t);
      root.append(si);
    });
    root.attrib.count = self.sharedStrings.length;
    root.attrib.uniqueCount = self.sharedStrings.length;
    self.archive.file(self.sharedStringsPath, etree.tostring(root));
  };
  Workbook.prototype.addSharedString = function(s) {
    var idx, self;
    self = this;
    idx = self.sharedStrings.length;
    self.sharedStrings.push(s);
    self.sharedStringsLookup[s] = idx;
    return idx;
  };
  Workbook.prototype.stringIndex = function(s) {
    var idx, self;
    self = this;
    idx = self.sharedStringsLookup[s];
    if (idx === void 0) {
      idx = self.addSharedString(s);
    }
    return idx;
  };
  Workbook.prototype.replaceString = function(oldString, newString) {
    var idx, self;
    self = this;
    idx = self.sharedStringsLookup[oldString];
    if (idx === void 0) {
      idx = self.addSharedString(newString);
    } else {
      self.sharedStrings[idx] = newString;
      delete self.sharedStringsLookup[oldString];
      self.sharedStringsLookup[newString] = idx;
    }
    return idx;
  };
  Workbook.prototype.loadSheets = function(prefix, workbook, workbookRels) {
    var self, sheets;
    self = this;
    sheets = [];
    workbook.findall('sheets/sheet').forEach(function(sheet) {
      var filename, relId, relationship, sheetId;
      sheetId = sheet.attrib.sheetId;
      relId = sheet.attrib['r:id'];
      relationship = workbookRels.find('Relationship[@Id=\'' + relId + '\']');
      filename = path.join(prefix, relationship.attrib.Target);
      sheets.push({
        id: parseInt(sheetId, 10),
        name: sheet.attrib.name,
        filename: filename
      });
    });
    return sheets;
  };
  Workbook.prototype.loadSheet = function(sheet) {
    var i, info, self;
    self = this;
    info = null;
    i = 0;
    while (i < self.sheets.length) {
      if (typeof sheet === 'number' && self.sheets[i].id === sheet || self.sheets[i].name === sheet) {
        info = self.sheets[i];
        break;
      }
      ++i;
    }
    if (info === null) {
      throw new Error('Sheet ' + sheet + ' not found');
    }
    info.filename = info.filename.replace(/\\/g, '/');
    return {
      filename: info.filename,
      name: info.name,
      id: info.id,
      root: etree.parse(self.archive.file(info.filename).asText()).getroot()
    };
  };
  Workbook.prototype.loadTables = function(sheet, sheetFilename) {
    var rels, relsFile, relsFilename, self, sheetDirectory, sheetName, tables;
    self = this;
    sheetDirectory = path.dirname(sheetFilename);
    sheetName = path.basename(sheetFilename);
    relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels');
    relsFile = self.archive.file(relsFilename);
    tables = [];
    if (relsFile === null) {
      return tables;
    }
    rels = etree.parse(relsFile.asText()).getroot();
    sheet.findall('tableParts/tablePart').forEach(function(tablePart) {
      var relationshipId, tableFilename, tableTree, target;
      relationshipId = tablePart.attrib['r:id'];
      target = rels.find('Relationship[@Id=\'' + relationshipId + '\']').attrib.Target;
      tableFilename = path.join(sheetDirectory, target);
      tableTree = etree.parse(self.archive.file(tableFilename).asText());
      tables.push({
        filename: tableFilename,
        root: tableTree.getroot()
      });
    });
    return tables;
  };
  Workbook.prototype.writeTables = function(tables) {
    var self;
    self = this;
    tables.forEach(function(namedTable) {
      self.archive.file(namedTable.filename, etree.tostring(namedTable.root));
    });
  };
  Workbook.prototype.substituteTableColumnHeaders = function(tables, substitutions) {
    var self;
    self = this;
    tables.forEach(function(table) {
      var autoFilter, columns, idx, inserted, newColumns, root, tableRange;
      root = table.root;
      columns = root.find('tableColumns');
      autoFilter = root.find('autoFilter');
      tableRange = self.splitRange(root.attrib.ref);
      idx = 0;
      inserted = 0;
      newColumns = [];
      columns.findall('tableColumn').forEach(function(col) {
        var name;
        ++idx;
        col.attrib.id = Number(idx).toString();
        newColumns.push(col);
        name = col.attrib.name;
        self.extractPlaceholders(name).forEach(function(placeholder) {
          var substitution;
          substitution = substitutions[placeholder.name];
          if (substitution === void 0) {
            return;
          }
          if (placeholder.full && placeholder.type === 'normal' && substitution instanceof Array) {
            substitution.forEach(function(element, i) {
              var newCol;
              newCol = col;
              if (i > 0) {
                newCol = self.cloneElement(newCol);
                newCol.attrib.id = Number(++idx).toString();
                newColumns.push(newCol);
                ++inserted;
                tableRange.end = self.nextCol(tableRange.end);
              }
              newCol.attrib.name = self.stringify(element);
            });
          } else {
            name = name.replace(placeholder.placeholder, self.stringify(substitution));
            col.attrib.name = name;
          }
        });
      });
      self.replaceChildren(columns, newColumns);
      if (inserted > 0) {
        columns.attrib.count = Number(idx).toString();
        root.attrib.ref = self.joinRange(tableRange);
        if (autoFilter !== null) {
          autoFilter.attrib.ref = self.joinRange(tableRange);
        }
      }
    });
  };
  Workbook.prototype.extractPlaceholders = function(string) {
    var match, matches, re;
    re = /\{{(?:(.+?):)?(.+?)(?:\.(.+?))?}}/g;
    match = null;
    matches = [];
    while ((match = re.exec(string)) !== null) {
      matches.push({
        placeholder: match[0],
        type: match[1] || 'normal',
        name: match[2],
        key: match[3],
        full: match[0].length === string.length
      });
    }
    return matches;
  };
  Workbook.prototype.splitRef = function(ref) {
    var match;
    match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/);
    return {
      table: match[1] || null,
      colAbsolute: Boolean(match[2]),
      col: match[3],
      rowAbsolute: Boolean(match[4]),
      row: parseInt(match[5], 10)
    };
  };
  Workbook.prototype.joinRef = function(ref) {
    return (ref.table ? ref.table + '!' : '') + (ref.colAbsolute ? '$' : '') + ref.col.toUpperCase() + (ref.rowAbsolute ? '$' : '') + Number(ref.row).toString();
  };
  Workbook.prototype.nextCol = function(ref) {
    var self;
    self = this;
    ref = ref.toUpperCase();
    return ref.replace(/[A-Z]+/, function(match) {
      return self.numToChar(self.charToNum(match) + 1);
    });
  };
  Workbook.prototype.nextRow = function(ref) {
    ref = ref.toUpperCase();
    return ref.replace(/[0-9]+/, function(match) {
      return (parseInt(match, 10) + 1).toString();
    });
  };
  Workbook.prototype.charToNum = function(str) {
    var idx, iteration, multiplier, num, thisChar;
    num = 0;
    idx = str.length - 1;
    iteration = 0;
    while (idx >= 0) {
      thisChar = str.charCodeAt(idx) - 64;
      multiplier = Math.pow(26, iteration);
      num += multiplier * thisChar;
      --idx;
      ++iteration;
    }
    return num;
  };
  Workbook.prototype.numToChar = function(num) {
    var charCode, i, remainder, str;
    str = '';
    i = 0;
    while (num > 0) {
      remainder = num % 26;
      charCode = remainder + 64;
      num = (num - remainder) / 26;
      if (remainder === 0) {
        charCode = 90;
        --num;
      }
      str = String.fromCharCode(charCode) + str;
      ++i;
    }
    return str;
  };
  Workbook.prototype.isRange = function(ref) {
    return ref.indexOf(':') !== -1;
  };
  Workbook.prototype.isWithin = function(ref, startRef, endRef) {
    var end, self, start, target;
    self = this;
    start = self.splitRef(startRef);
    end = self.splitRef(endRef);
    target = self.splitRef(ref);
    start.col = self.charToNum(start.col);
    end.col = self.charToNum(end.col);
    target.col = self.charToNum(target.col);
    return start.row <= target.row && target.row <= end.row && start.col <= target.col && target.col <= end.col;
  };
  Workbook.prototype.stringify = function(value) {
    var self;
    self = this;
    if (value instanceof Date) {
      return value.toISOString();
    } else if (typeof value === 'number' || typeof value === 'boolean') {
      return Number(value).toString();
    } else {
      return String(value).toString();
    }
  };
  Workbook.prototype.insertCellValue = function(cell, substitution) {
    var cellValue, self, stringified;
    self = this;
    cellValue = cell.find('v');
    stringified = self.stringify(substitution);
    if (typeof substitution === 'number') {
      cell.attrib.t = 'n';
      cellValue.text = stringified;
    } else if (typeof substitution === 'boolean') {
      cell.attrib.t = 'b';
      cellValue.text = stringified;
    } else if (substitution instanceof Date) {
      cell.attrib.t = 'd';
      cellValue.text = stringified;
    } else {
      cell.attrib.t = 's';
      cellValue.text = Number(self.stringIndex(stringified)).toString();
    }
    return stringified;
  };
  Workbook.prototype.stripValue = function(cell) {
    return cell._children = cell._children.remove(function(child) {
      return child.tag === 'v';
    });
  };
  Workbook.prototype.substituteScalar = function(cell, string, placeholder, substitution) {
    var newString, self;
    self = this;
    if (placeholder.full && typeof substitution === 'string') {
      self.replaceString(string, substitution);
    }
    if (placeholder.full) {
      return self.insertCellValue(cell, substitution);
    } else {
      newString = string.replace(placeholder.placeholder, self.stringify(substitution));
      cell.attrib.t = 's';
      self.replaceString(string, newString);
      return newString;
    }
  };
  Workbook.prototype.substituteArray = function(cells, cell, substitution) {
    var currentCell, newCellsInserted, self;
    self = this;
    newCellsInserted = -1;
    currentCell = cell.attrib.r;
    substitution.forEach(function(element) {
      var newCell;
      ++newCellsInserted;
      if (newCellsInserted > 0) {
        currentCell = self.nextCol(currentCell);
      }
      newCell = self.cloneElement(cell);
      self.insertCellValue(newCell, element);
      newCell.attrib.r = currentCell;
      cells.push(newCell);
    });
    return newCellsInserted;
  };
  Workbook.prototype.substituteTable = function(row, newTableRows, cells, cell, namedTables, substitution, key) {
    var newCellsInserted, parentTables, self;
    self = this;
    newCellsInserted = 0;
    if (substitution.length === 0) {
      delete cell.attrib.t;
      self.replaceChildren(cell, []);
    } else {
      parentTables = namedTables.filter(function(namedTable) {
        var range;
        range = self.splitRange(namedTable.root.attrib.ref);
        return self.isWithin(cell.attrib.r, range.start, range.end);
      });
      substitution.forEach(function(element, idx) {
        var newCell, newCells, newCellsInsertedOnNewRow, newRow, value;
        newRow = void 0;
        newCell = void 0;
        newCellsInsertedOnNewRow = 0;
        newCells = [];
        value = element[key];
        if (idx === 0) {
          if (value instanceof Array) {
            newCellsInserted = self.substituteArray(cells, cell, value);
          } else {
            self.insertCellValue(cell, value);
          }
        } else {
          if (idx - 1 < newTableRows.length) {
            newRow = newTableRows[idx - 1];
          } else {
            newRow = self.cloneElement(row, false);
            newRow.attrib.r = self.getCurrentRow(row, newTableRows.length + 1);
            newTableRows.push(newRow);
          }
          newCell = self.cloneElement(cell);
          newCell.attrib.r = self.joinRef({
            row: newRow.attrib.r,
            col: self.splitRef(newCell.attrib.r).col
          });
          if (value instanceof Array) {
            newCellsInsertedOnNewRow = self.substituteArray(newCells, newCell, value);
            newCells.forEach(function(newCell) {
              newRow.append(newCell);
            });
            self.updateRowSpan(newRow, newCellsInsertedOnNewRow);
          } else {
            self.insertCellValue(newCell, value);
            newRow.append(newCell);
          }
          parentTables.forEach(function(namedTable) {
            var autoFilter, range, tableRoot;
            tableRoot = namedTable.root;
            autoFilter = tableRoot.find('autoFilter');
            range = self.splitRange(tableRoot.attrib.ref);
            if (!self.isWithin(newCell.attrib.r, range.start, range.end)) {
              range.end = self.nextRow(range.end);
              tableRoot.attrib.ref = self.joinRange(range);
              if (autoFilter !== null) {
                autoFilter.attrib.ref = tableRoot.attrib.ref;
              }
            }
          });
        }
      });
    }
    return newCellsInserted;
  };
  Workbook.prototype.cloneElement = function(element, deep) {
    var newElement, self;
    self = this;
    newElement = etree.Element(element.tag, element.attrib);
    newElement.text = element.text;
    newElement.tail = element.tail;
    if (deep !== false) {
      element.getchildren().forEach(function(child) {
        newElement.append(self.cloneElement(child, deep));
      });
    }
    return newElement;
  };
  Workbook.prototype.replaceChildren = function(parent, children) {
    parent.delSlice(0, parent.len());
    children.forEach(function(child) {
      parent.append(child);
    });
  };
  Workbook.prototype.getCurrentRow = function(row, rowsInserted) {
    return parseInt(row.attrib.r, 10) + rowsInserted;
  };
  Workbook.prototype.getCurrentCell = function(cell, currentRow, cellsInserted) {
    var colNum, colRef, self;
    self = this;
    colRef = self.splitRef(cell.attrib.r).col;
    colNum = self.charToNum(colRef);
    return self.joinRef({
      row: currentRow,
      col: self.numToChar(colNum + cellsInserted)
    });
  };
  Workbook.prototype.updateRowSpan = function(row, cellsInserted) {
    var rowSpan;
    if (cellsInserted !== 0 && row.attrib.spans) {
      rowSpan = row.attrib.spans.split(':').map(function(f) {
        return parseInt(f, 10);
      });
      rowSpan[1] += cellsInserted;
      row.attrib.spans = rowSpan.join(':');
    }
  };
  Workbook.prototype.splitRange = function(range) {
    var split;
    split = range.split(':');
    return {
      start: split[0],
      end: split[1]
    };
  };
  Workbook.prototype.joinRange = function(range) {
    return range.start + ':' + range.end;
  };
  Workbook.prototype.pushRight = function(workbook, sheet, currentCell, numCols) {
    var cellRef, currentCol, currentRow, self;
    self = this;
    cellRef = self.splitRef(currentCell);
    currentRow = cellRef.row;
    currentCol = self.charToNum(cellRef.col);
    sheet.findall('mergeCells/mergeCell').forEach(function(mergeCell) {
      var mergeEnd, mergeEndCol, mergeRange, mergeStart, mergeStartCol;
      mergeRange = self.splitRange(mergeCell.attrib.ref);
      mergeStart = self.splitRef(mergeRange.start);
      mergeStartCol = self.charToNum(mergeStart.col);
      mergeEnd = self.splitRef(mergeRange.end);
      mergeEndCol = self.charToNum(mergeEnd.col);
      if (mergeStart.row === currentRow && currentCol < mergeStartCol) {
        mergeStart.col = self.numToChar(mergeStartCol + numCols);
        mergeEnd.col = self.numToChar(mergeEndCol + numCols);
        mergeCell.attrib.ref = self.joinRange({
          start: self.joinRef(mergeStart),
          end: self.joinRef(mergeEnd)
        });
      }
    });
    workbook.findall('definedNames/definedName').forEach(function(name) {
      var namedCol, namedEnd, namedEndCol, namedRange, namedRef, namedStart, namedStartCol, ref;
      ref = name.text;
      if (self.isRange(ref)) {
        namedRange = self.splitRange(ref);
        namedStart = self.splitRef(namedRange.start);
        namedStartCol = self.charToNum(namedStart.col);
        namedEnd = self.splitRef(namedRange.end);
        namedEndCol = self.charToNum(namedEnd.col);
        if (namedStart.row === currentRow && currentCol < namedStartCol) {
          namedStart.col = self.numToChar(namedStartCol + numCols);
          namedEnd.col = self.numToChar(namedEndCol + numCols);
          name.text = self.joinRange({
            start: self.joinRef(namedStart),
            end: self.joinRef(namedEnd)
          });
        }
      } else {
        namedRef = self.splitRef(ref);
        namedCol = self.charToNum(namedRef.col);
        if (namedRef.row === currentRow && currentCol < namedCol) {
          namedRef.col = self.numToChar(namedCol + numCols);
          name.text = self.joinRef(namedRef);
        }
      }
    });
  };
  Workbook.prototype.pushDown = function(workbook, sheet, tables, currentRow, numRows) {
    var self;
    self = this;
    sheet.findall('mergeCells/mergeCell').forEach(function(mergeCell) {
      var mergeEnd, mergeRange, mergeStart;
      mergeRange = self.splitRange(mergeCell.attrib.ref);
      mergeStart = self.splitRef(mergeRange.start);
      mergeEnd = self.splitRef(mergeRange.end);
      if (mergeStart.row > currentRow) {
        mergeStart.row += numRows;
        mergeEnd.row += numRows;
        mergeCell.attrib.ref = self.joinRange({
          start: self.joinRef(mergeStart),
          end: self.joinRef(mergeEnd)
        });
      }
    });
    tables.forEach(function(table) {
      var autoFilter, tableEnd, tableRange, tableRoot, tableStart;
      tableRoot = table.root;
      tableRange = self.splitRange(tableRoot.attrib.ref);
      tableStart = self.splitRef(tableRange.start);
      tableEnd = self.splitRef(tableRange.end);
      if (tableStart.row > currentRow) {
        tableStart.row += numRows;
        tableEnd.row += numRows;
        tableRoot.attrib.ref = self.joinRange({
          start: self.joinRef(tableStart),
          end: self.joinRef(tableEnd)
        });
        autoFilter = tableRoot.find('autoFilter');
        if (autoFilter !== null) {
          autoFilter.attrib.ref = tableRoot.attrib.ref;
        }
      }
    });
    workbook.findall('definedNames/definedName').forEach(function(name) {
      var namedCol, namedEnd, namedRange, namedRef, namedStart, ref;
      ref = name.text;
      if (self.isRange(ref)) {
        namedRange = self.splitRange(ref);
        namedStart = self.splitRef(namedRange.start);
        namedEnd = self.splitRef(namedRange.end);
        if (namedStart.row > currentRow) {
          namedStart.row += numRows;
          namedEnd.row += numRows;
          name.text = self.joinRange({
            start: self.joinRef(namedStart),
            end: self.joinRef(namedEnd)
          });
        }
      } else {
        namedRef = self.splitRef(ref);
        namedCol = self.charToNum(namedRef.col);
        if (namedRef.row > currentRow) {
          namedRef.row += numRows;
          name.text = self.joinRef(namedRef);
        }
      }
    });
  };
  return Workbook;
})();
