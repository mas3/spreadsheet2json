# Spreadsheet to JSON

A command line program to convert a spreadsheet (*.xlsx file) to JSON.

## Requirement

- .NET 6

## Usage

```
spreadsheet2json [options]
```

## Options

```
-i, --input <filename> (REQUIRED)  Input xlsx filename.
-o, --output <filename>            Output json filename. If omitted, output JSON string to standard output.
--cell-format                      Include cell format data. [default: False]
--cell-row-column                  Include cell row and column address data. [default: False]
--debug                            Show debug message. [default: False]
--no-cell-data                     Suppress including cell data. [default: False]
--no-encode                        Suppress encoding process of JSON string. [default: False]
--no-indent                        Suppress indent process of JSON string. [default: False]
--no-properties                    Suppress including workbook property data. [default: False]
--object-format                    Represent sheet and cell data in object format. [default: False]
--sheet-info                       Include sheet information. [default: False]
--sheets <sheets>                  Target sheet names. Separate multiple sheets with colons. e.g., sheet1:sheet2
--version                          Show version information
-?, -h, --help                     Show help and usage information
```

## Sample

### Spreadsheet

|       | A         | B     |
| :-:   | :-:       | :-:   |
| **1** | A1        | B1    |
| **2** | 2023/1/23 | 12:34 |

### Example

```
spreadsheet2json --input example.xlsx
{
  "Properties": {
    "Path": "example.xlsx",
    "Author": "Name",
    "LastModifiedBy": "Name",
    "Company": ""
  },
  "Workbook": {
    "Sheets": [
      {
        "Name": "Sheet1",
        "Cells": [
          {
            "Address": "A1",
            "Text": "A1",
            "Type": "Text",
            "Value": "A1"
          },
          {
            "Address": "B1",
            "Text": "A2",
            "Type": "Text",
            "Value": "A2"
          },
          {
            "Address": "A2",
            "Text": "2023/1/23",
            "Type": "DateTime",
            "Value": 44949
          },
          {
            "Address": "B2",
            "Text": "12:34",
            "Type": "DateTime",
            "Value": 0.5236111111111111
          }
        ]
      }
    ]
  },
  "Tool": {
    "Name": "Spreadsheet to Json",
    "Version": "0.0.0"
  }
}
```

```
spreadsheet2json --input example.xlsx --object-format
{
  "Properties": {
    "Path": "example.xlsx",
    "Author": "Name",
    "LastModifiedBy": "Name",
    "Company": ""
  },
  "Workbook": {
    "Sheets": {
      "Sheet1": {
        "Name": "Sheet1",
        "Cells": {
          "A1": {
            "Address": "A1",
            "Text": "A1",
            "Type": "Text",
            "Value": "A1"
          },
          "B1": {
            "Address": "B1",
            "Text": "A2",
            "Type": "Text",
            "Value": "A2"
          },
          "A2": {
            "Address": "A2",
            "Text": "2023/1/23",
            "Type": "DateTime",
            "Value": 44949
          },
          "B2": {
            "Address": "B2",
            "Text": "12:34",
            "Type": "DateTime",
            "Value": 0.5236111111111111
          }
        }
      }
    }
  },
  "Tool": {
    "Name": "Spreadsheet to Json",
    "Version": "0.0.0"
  }
}
```
