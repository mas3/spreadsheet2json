using System.Globalization;
using System.Text.Json.Nodes;

namespace Spreadsheet2Json.tests
{
    public class SpreadsheetTranslatorTests
    {
        private readonly string _fileSheetInformation = @"testdata\sheetinformation.xlsx";

        [Fact]
        public void SheetList()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            Assert.Equal("visible", sheets[0]?.AsObject()["Name"]?.ToString());
            Assert.Equal("hidden", sheets[1]?.AsObject()["Name"]?.ToString());
            Assert.Equal("veryhidden", sheets[2]?.AsObject()["Name"]?.ToString());
            Assert.Equal("protect", sheets[3]?.AsObject()["Name"]?.ToString());
            Assert.Equal("protect_password", sheets[4]?.AsObject()["Name"]?.ToString());
            Assert.Equal("active", sheets[5]?.AsObject()["Name"]?.ToString());
            Assert.Equal("selected", sheets[6]?.AsObject()["Name"]?.ToString());
        }

        [Fact]
        public void SuppressSheetInformation()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeSheetInformation = false
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("visible") ?? false);
            Assert.NotNull(sheet);

            Assert.Null(sheet["Visible"]);
        }

        [Fact]
        public void SheetVisibility()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeSheetInformation = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("visible") ?? false);
            Assert.NotNull(sheet);
            Assert.True(sheet["Visible"]?.GetValue<bool>());
            Assert.Equal("Visible", sheet["Visibility"]?.ToString());

            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("hidden") ?? false);
            Assert.NotNull(sheet);
            Assert.False(sheet["Visible"]?.GetValue<bool>());
            Assert.Equal("Hidden", sheet["Visibility"]?.ToString());

            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("veryhidden") ?? false);
            Assert.NotNull(sheet);
            Assert.False(sheet["Visible"]?.GetValue<bool>());
            Assert.Equal("VeryHidden", sheet["Visibility"]?.ToString());
        }

        [Fact]
        public void SheetProtection()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeSheetInformation = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("protect") ?? false);
            Assert.NotNull(sheet);
            Assert.True(sheet["Protected"]?.GetValue<bool>());
            Assert.False(sheet["PasswordProtected"]?.GetValue<bool>());

            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("protect_password") ?? false);
            Assert.NotNull(sheet);
            Assert.True(sheet["Protected"]?.GetValue<bool>());
            Assert.True(sheet["PasswordProtected"]?.GetValue<bool>());
        }

        [Fact]
        public void SheetSelection()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeSheetInformation = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("active") ?? false);
            Assert.NotNull(sheet);
            Assert.True(sheet["Active"]?.GetValue<bool>());
            Assert.True(sheet["Selected"]?.GetValue<bool>());

            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("selected") ?? false);
            Assert.NotNull(sheet);
            Assert.False(sheet["Active"]?.GetValue<bool>());
            Assert.True(sheet["Selected"]?.GetValue<bool>());
        }

        [Fact]
        public void IncludeCellFormat()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeCellFormat = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("visible") ?? false);
            Assert.NotNull(sheet);

            JsonNode? cells = sheet["Cells"];
            Assert.NotNull(cells);

            JsonNode? a1 = cells.AsArray().First(item => item?.AsObject()["Address"]?.ToString().Equals("A1") ?? false);
            Assert.NotNull(a1);

            Assert.NotNull(a1["NumberFormat_Format"]);
            Assert.NotNull(a1["NumberFormat_Id"]);
        }

        [Fact]
        public void IncludeCellRowColumn()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeCellRowColumn = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("visible") ?? false);
            Assert.NotNull(sheet);

            JsonNode? cells = sheet["Cells"];
            Assert.NotNull(cells);

            JsonNode? a1 = cells.AsArray().First(item => item?.AsObject()["Address"]?.ToString().Equals("A1") ?? false);
            Assert.NotNull(a1);

            Assert.NotNull(a1["Row"]);
            Assert.NotNull(a1["Column"]);
        }

        [Fact]
        public void NotIncludeCelldata()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeCellData = false
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("visible") ?? false);
            Assert.NotNull(sheet);

            JsonNode? cells = sheet["Cells"];
            Assert.Null(cells);
        }

        [Fact]
        public void NoIndent()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                Indented = false
            };

            string jsonString = st.GetJsonString();
            Assert.Equal(-1, jsonString.IndexOf("\n"));
        }

        [Fact]
        public void NotIncludeProperties()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeProperties = false
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonNode? properties = jsonNode["Properties"];
            Assert.Null(properties);
        }

        [Fact]
        public void ObjectFormat()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsObjectFormat = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonNode? sheet = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsObject()["visible"];
            Assert.NotNull(sheet);

            JsonNode? cell = sheet["Cells"]?.AsObject()["A1"];
            Assert.NotNull(cell);
        }

        [Fact]
        public void IndicateSheets()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                SheetNames = "visible:hidden"
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            Assert.Equal(2, sheets.Count);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("visible") ?? false);
            Assert.NotNull(sheet);

            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("hidden") ?? false);
            Assert.NotNull(sheet);

            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("veryhidden") ?? false);
            Assert.Null(sheet);
        }

        [Fact]
        public void SheetZoomScale()
        {
            SpreadsheetTranslator st = new(CultureInfo.CurrentCulture.Name)
            {
                InputFilePath = _fileSheetInformation,
                IsIncludeSheetInformation = true
            };

            string jsonString = st.GetJsonString();
            JsonNode? jsonNode = JsonNode.Parse(jsonString);
            Assert.NotNull(jsonNode);

            JsonArray? sheets = jsonNode["Workbook"]?.AsObject()["Sheets"]?.AsArray();
            Assert.NotNull(sheets);

            JsonNode? sheet;
            sheet = sheets.FirstOrDefault(item => item?.AsObject()["Name"]?.ToString().Equals("zoom150") ?? false);
            Assert.NotNull(sheet);
            Assert.Equal(150, sheet["ZoomScale"]?.GetValue<int>());
        }

    }
}
