# xlsx-injector

This module allows injection of values into an XLSX template with styling using angular-styled tags {{}}

It is heavily based on Martin Aspeli's (https://github.com/optilude) `xlsx-template` , but is modified in an 
opinionated fashion to work better in certain workflows. If you are looking for more flexibility and/or control, 
you may be better off using the `xlsx-template` (https://github.com/optilude/xlsx-template) library instead.

## What It Does

This module allows you to use any existing `xlsx` file and inject values into it without losing any 
styling information.

## Why Would You Use This

If you're looking to create dynamic Excel based reports, this tool will allow you to do that.

## How To Install This

`npm install xlsx-injector`

## How To Use

### Basic Workflow

* Create a template file as an `.xlsx`
* Style it up, and use placeholders
* Using the `xlsx-injector`, load the template `xlsx`
* Inject substitution values via a Javascript object
^ Save the file as a new `.xlsx`

### Placeholders

Placeholders are inserted in cells in a spreadsheet. It does not matter how
those cells are formatted, so e.g. it is OK to insert a placeholder (which is
text content) into a cell formatted as a number or currecy or date, if you
expect the placeholder to resolve to a number or currency or date.

This module uses {{ }} placeholders, which should be familiar to users of Handlebars or Angular.

#### Scalars

Simple placholders take the format `{{name}}`. Here, `name` is the name of a
key in the placeholders map. The value of this placholder here should be a
scalar, i.e. not an array or object. The placeholder may appear on its own in a
cell, or as part of a text string. For example:

    | Extracted on: | {{extractDate}} |

might result in (depending on date formatting in the second cell):

    | Extracted on: | Jun-01-2013 |

Here, `extractDate` may be a date and the second cell may be formatted as a
number.

#### Columns

You can use arrays as placeholder values to indicate that the placeholder cell
is to be replicated across columns. In this case, the placeholder cannot appear
inside a text string - it must be the only thing in its cell. For example,
if the placehodler value `dates` is an array of dates:

    | {{dates}} |

might result in:

    | Jun-01-2013 | Jun-02-2013 | Jun-03-2013 |

#### Tables

Finally, you can build tables made up of multiple rows. In this case, each
placeholder should be prefixed by `table:` and contain both the name of the
placeholder variable (a list of objects) and a key (in each object in the list).
For example:

    | Name                  | Age                  |
    | {{table:people.name}} | {{table:people.age}} |

If the replacement value under `people` is an array of objects, and each of
those objects have keys `name` and `age`, you may end up with something like:

    | Name        | Age |
    | John Smith  | 20  |
    | Bob Johnson | 22  |

If a particular value is an array, then it will be repeated accross columns as
above.

### Generating reports

To make this magic happen, you need some code like this:

(Coffeescript example)
```
  XlsxInjector = require('xlsx-injector')

  templatePath = "./path/to/template.xlsx"
  outputPath = "./path/to/output_file.xlsx"

  # Object containing attributes that match the placeholder tokens in the template
  values = 
    extractDate: new Date()
    dates: [new Date("2013-06-01"), new Date("2013-06-02"), new Date("2013-06-03")]
    people: [
        {name: "John Smith", age: 20}
        {name: "Bob Johnson", age: 22}
    ]

  # Open a workbook
  workbook = new XlsxInjector(templatePath)
  sheetNumber = 1
  workbook.substitute sheetNumber, values
  # Save the workbook
  workbook.writeFile(outputPath)

```


## Caveats

* Only `.xlsx` is supported
* Column (array) and table (array-of-objects) insertions cause rows and cells to
  be inserted or removed. When this happens, only a limited number of
  adjustments are made:
    * Merged cells and named cells/ranges to the right of cells where insertions
      or deletions are made are moved right or left, appropriately. This may
      not work well if cells are merged across rows, unless all rows have the
      same number of insertions.
    * Merged cells, named tables or named cells/ranges below rows where further
      rows are inserted are moved down.
  Formulae are not adjusted.
* Placeholders only work in simple cells and tables, pivot tables or
  other such things.

## Things At Risk of Changing

* The `table:` prefix is subject to change in a newer version

## Changelog

### 0.0.1

* Initial release
* Convert original code to Coffeescript
* Add Windows compatibility
* Purge elements on cells with `#VALUE!` to allow for auto-calculation of formula cells when the output workbook is loaded in Excel or Office 365.
* Bake in XLSX load and write functions

## Attributions

Functionality not in the original `xlsx-template` is Copyright (c) 2015, Assign It To Us Technologies

Original `xlsx-template` code is Copyright (c) 2013 Martin Aspeli

If you are looking for the xlsx-template npm, it is available here: https://github.com/optilude/xlsx-template