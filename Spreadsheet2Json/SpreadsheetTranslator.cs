using System;
using System.Collections.Generic;
using System.CommandLine;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Unicode;
using ClosedXML.Excel;

namespace Spreadsheet2Json
{
    /// <summary>
    /// Spreadsheet Translator class.
    /// </summary>
    internal class SpreadsheetTranslator
    {
        private readonly CellData _cellData;

        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetTranslator"/> class.
        /// </summary>
        /// <param name="cultureName">Culture name.</param>
        public SpreadsheetTranslator(string cultureName)
        {
            _cellData = new CellData(cultureName);

            Encoded = false;
            IsIncludeCellData = true;
            IsIncludeCellFormat = false;
            IsIncludeCellRowColumn = false;
            IsIncludeSheetInformation = false;
            IsIncludeProperties = true;
            IsObjectFormat = false;
            Indented = true;
            InputFilePath = string.Empty;
            SheetNames = string.Empty;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the json is encoded.
        /// </summary>
        public bool Encoded { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether json is indented.
        /// </summary>
        public bool Indented { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the include cell data to json.
        /// </summary>
        public bool IsIncludeCellData { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the include cell format to json.
        /// </summary>
        public bool IsIncludeCellFormat { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the include cell row and column address to json.
        /// </summary>
        public bool IsIncludeCellRowColumn { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the include worksheet property to json.
        /// </summary>
        public bool IsIncludeProperties { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the include sheet information to json.
        /// </summary>
        public bool IsIncludeSheetInformation { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the represent sheet and cell data in object format.
        /// </summary>
        public bool IsObjectFormat { get; set; }

        /// <summary>
        /// Gets or sets the input filepath.
        /// </summary>
        public string InputFilePath { get; set; }

        /// <summary>
        /// Gets or sets the sheet names.
        /// </summary>
        /// <remarks>
        /// If there are multiple sheet names, separate them with colons.
        /// </remarks>
        /// <example>e.g., "sheet1", "sheet1:sheet2".</example>
        public string SheetNames { get; set; }

        /// <summary>
        /// Get a JSON-formatted spreadsheet data string.
        /// </summary>
        /// <returns>Spreadsheet data string in JSON format.</returns>
        public string GetJsonString()
        {
            if (InputFilePath == string.Empty)
            {
                throw new InvalidOperationException("input filepath empty.");
            }

            if (!File.Exists(InputFilePath))
            {
                throw new InvalidOperationException($"{InputFilePath} does not exist.");
            }

            using FileStream fs = new(InputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using XLWorkbook workbook = new(fs);

            JsonObject jsonWorkbook = new();
            List<string> sheetNameList = SheetNames.Split(':').Select(s => s.Trim()).ToList();
            if (IsObjectFormat)
            {
                JsonObject sheets = new();
                foreach (IXLWorksheet? worksheet in workbook.Worksheets)
                {
                    if (SheetNames != string.Empty && !sheetNameList.Any(s => string.Equals(s, worksheet.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        // Skip unspecified sheets.
                        continue;
                    }

                    JsonObject sheet = GetJsonSheetData(worksheet);
                    sheets[worksheet.Name] = sheet;
                }

                jsonWorkbook["Sheets"] = sheets;
            }
            else
            {
                JsonArray sheets = new();
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (SheetNames != string.Empty && !sheetNameList.Any(s => string.Equals(s, worksheet.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        // Skip unspecified sheets.
                        continue;
                    }

                    JsonObject sheet = GetJsonSheetData(worksheet);
                    sheets.Add(sheet);
                }

                jsonWorkbook["Sheets"] = sheets;
            }

            JsonObject data = new();

            if (IsIncludeProperties)
            {
                data["Properties"] = GetJsonPropertyData(workbook);
            }

            data["Workbook"] = jsonWorkbook;
            data["Tool"] = GetJsonToolData();

            JsonSerializerOptions option = new()
            {
                WriteIndented = Indented,
            };
            if (!Encoded)
            {
                option.Encoder = JavaScriptEncoder.Create(UnicodeRanges.All);
            }

            return data.ToJsonString(option);
        }

        /// <summary>
        /// Make property data for JSON.
        /// </summary>
        /// <param name="workbook">Workbook object.</param>
        /// <returns>Property data for JSON.</returns>
        private JsonObject GetJsonPropertyData(XLWorkbook workbook)
        {
            return new JsonObject()
            {
                ["Path"] = InputFilePath,
                ["Author"] = workbook.Properties.Author,
                ["LastModifiedBy"] = workbook.Properties.LastModifiedBy,
                ["Company"] = workbook.Properties.Company,
            };
        }

        /// <summary>
        /// Make tool data for JSON.
        /// </summary>
        /// <returns>Tool data for JSON.</returns>
        private JsonObject GetJsonToolData()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();

            return new JsonObject()
            {
                ["Name"] = "Spreadsheet to Json",
                ["Version"] = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
            };
        }

        /// <summary>
        /// Make sheet data for JSON.
        /// </summary>
        /// <param name="worksheet">Worksheet object.</param>
        /// <returns>Sheet data for JSON.</returns>
        private JsonObject GetJsonSheetData(IXLWorksheet worksheet)
        {
            JsonObject sheet = new()
            {
                ["Name"] = worksheet.Name,
            };

            if (IsIncludeSheetInformation)
            {
                sheet["Protected"] = worksheet.IsProtected;
                sheet["PasswordProtected"] = worksheet.IsPasswordProtected;
                sheet["Visible"] = worksheet.Visibility == XLWorksheetVisibility.Visible;
                sheet["Visibility"] = worksheet.Visibility.ToString();
                sheet["Active"] = worksheet.TabActive;
                sheet["Selected"] = worksheet.TabSelected;
            }

            if (IsIncludeCellData)
            {
                if (IsObjectFormat)
                {
                    JsonObject cells = new();
                    foreach (IXLCell? cell in worksheet.CellsUsed())
                    {
                        string? address = cell.Address.ToString();
                        if (address != null)
                        {
                            cells[address] = GetJsonCellData(cell);
                        }
                    }

                    sheet["Cells"] = cells;
                }
                else
                {
                    JsonArray cells = new();
                    foreach (IXLCell? cell in worksheet.CellsUsed())
                    {
                        cells.Add(GetJsonCellData(cell));
                    }

                    sheet["Cells"] = cells;
                }
            }

            return sheet;
        }

        /// <summary>
        /// Make cell data for JSON.
        /// </summary>
        /// <param name="cell">Spread sheet cell object.</param>
        /// <returns>Cell data for JSON.</returns>
        private JsonObject GetJsonCellData(IXLCell cell)
        {
            JsonObject jsonCellData = new();
            _cellData.Cell = cell;

            jsonCellData["Address"] = cell.Address.ToString();

            if (IsIncludeCellRowColumn)
            {
                jsonCellData["Row"] = cell.Address.RowNumber;
                jsonCellData["Column"] = cell.Address.ColumnNumber;
            }

            jsonCellData["Text"] = _cellData.Text;
            jsonCellData["Type"] = _cellData.Type.ToString();
            if (cell.FormulaA1 != string.Empty)
            {
                jsonCellData["Formula"] = cell.FormulaA1;
            }

            jsonCellData["Value"] = _cellData.Type switch
            {
                XLDataType.DateTime => (JsonNode)cell.Value.GetUnifiedNumber(),
                XLDataType.Number => (JsonNode)cell.Value.GetNumber(),
                _ => (JsonNode?)cell.Value.ToString(),
            };
            string memo = cell.GetComment().Text;
            if (memo != string.Empty)
            {
                jsonCellData["Memo"] = cell.GetComment().Text;
            }

            if (IsIncludeCellFormat)
            {
                jsonCellData["NumberFormat_Format"] = cell.Style.NumberFormat.Format;
                jsonCellData["NumberFormat_Id"] = cell.Style.NumberFormat.NumberFormatId;
            }

            return jsonCellData;
        }
    }
}
