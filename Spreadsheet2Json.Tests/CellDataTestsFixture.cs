using ClosedXML.Excel;

namespace Spreadsheet2Json.tests
{
    public class CellDataTestsFixture : IDisposable
    {
        private readonly XLWorkbook _workbookDateTime;

        public XLWorkbook WorkbookDateTime
        {
            get { return _workbookDateTime; }
        }

        public CellDataTestsFixture()
        {
            string filepath = @"testdata\datetime.xlsx";
            using FileStream fs = new(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _workbookDateTime = new XLWorkbook(fs);
        }

        public void Dispose()
        {
            _workbookDateTime.Dispose();
            GC.SuppressFinalize(this);
        }

        ~CellDataTestsFixture()
        {
            _workbookDateTime.Dispose();
        }
    }
}
