using ClosedXML.Excel;
using System.Globalization;

namespace Spreadsheet2Json.tests
{
    public class CellDataTests : IClassFixture<CellDataTestsFixture>
    {
        readonly CellDataTestsFixture _fixture;

        public CellDataTests(CellDataTestsFixture fixture)
        {
            _fixture = fixture;
        }

        [SkippableFact]
        public void DateTime_ShortDate_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A1");
            cellData.Cell = cell;

            Assert.Equal("2023/3/9", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [SkippableFact]
        public void DateTime_LongDate_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A2");
            cellData.Cell = cell;

            Assert.Equal("2023年3月9日", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_yyyy_mm_dd()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A3");
            cellData.Cell = cell;

            Assert.Equal("2023-03-09", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Japanese_yyyy_mm_dd()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A4");
            cellData.Cell = cell;

            Assert.Equal("2023年3月9日", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Japanese_yyyy_mm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A5");
            cellData.Cell = cell;

            Assert.Equal("2023年3月", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Japanese_mm_dd()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A6");
            cellData.Cell = cell;

            Assert.Equal("3月9日", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_yyyy_m_d()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A7");
            cellData.Cell = cell;

            Assert.Equal("2023/3/9", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_yyyy_m_d_h_mm_AMPM()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A8");
            cellData.Cell = cell;

            Assert.Equal("2023/3/9 12:00 AM", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_yyyy_m_d_h_mm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A9");
            cellData.Cell = cell;

            Assert.Equal("2023/3/9 0:00", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_m_d()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A10");
            cellData.Cell = cell;

            Assert.Equal("3/9", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_m_d_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A11");
            cellData.Cell = cell;

            Assert.Equal("3/9/23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_mm_dd_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A12");
            cellData.Cell = cell;

            Assert.Equal("03/09/23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_d_mmm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A13");
            cellData.Cell = cell;

            Assert.Equal("9-Mar", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_d_mmm_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A14");
            cellData.Cell = cell;

            Assert.Equal("9-Mar-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_dd_mmm_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A15");
            cellData.Cell = cell;

            Assert.Equal("09-Mar-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_mmm_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A16");
            cellData.Cell = cell;

            Assert.Equal("Mar-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_mmmm_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A17");
            cellData.Cell = cell;

            Assert.Equal("March-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_mmmmm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A18");
            cellData.Cell = cell;

            Assert.Equal("M", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_mmmmm_yy()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A19");
            cellData.Cell = cell;

            Assert.Equal("M-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [SkippableFact]
        public void DateTime_Id31_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A20");
            cellData.Cell = cell;

            Assert.Equal("2023年3月9日", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [SkippableFact]
        public void DateTime_Id55_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A21");
            cellData.Cell = cell;

            Assert.Equal("2023年3月", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [SkippableFact]
        public void DateTime_Id56_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A22");
            cellData.Cell = cell;

            Assert.Equal("3月9日", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [SkippableFact]
        public void DateTime_Id30_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A23");
            cellData.Cell = cell;

            Assert.Equal("3/9/23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Id15()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A24");
            cellData.Cell = cell;

            Assert.Equal("9-Mar-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Id16()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A25");
            cellData.Cell = cell;

            Assert.Equal("9-Mar", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Id17()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("date_ja-JP").Cell("A26");
            cellData.Cell = cell;

            Assert.Equal("Mar-23", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [SkippableFact]
        public void DateTime_LongTime_jaJP()
        {
            Skip.IfNot(CultureInfo.CurrentCulture.Name == "ja-JP", "Test only ja.JP culture.");

            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A1");
            cellData.Cell = cell;

            Assert.Equal("4:05:06", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_h_mm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A2");
            cellData.Cell = cell;

            Assert.Equal("4:05", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_h_mm_ampm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A3");
            cellData.Cell = cell;

            Assert.Equal("4:05 AM", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_h_mm_ss()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A4");
            cellData.Cell = cell;

            Assert.Equal("4:05:06", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_h_mm_ss_ampm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A5");
            cellData.Cell = cell;

            Assert.Equal("4:05:06 AM", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_US_yyyy_m_d__h_mm_ampm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A6");
            cellData.Cell = cell;

            Assert.Equal("1900/1/1 4:05 AM", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_yyyy_m_d__h_mm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A7");
            cellData.Cell = cell;

            Assert.Equal("1900/1/1 4:05", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Japanese_h_mm()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A8");
            cellData.Cell = cell;

            Assert.Equal("4時05分", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }

        [Fact]
        public void DateTime_Japanese_h_mm_ss()
        {
            CellData cellData = new(CultureInfo.CurrentCulture.Name);
            IXLCell cell = _fixture.WorkbookDateTime.Worksheet("time_ja-JP").Cell("A9");
            cellData.Cell = cell;

            Assert.Equal("4時05分06秒", cellData.Text);
            Assert.Equal(XLDataType.DateTime, cellData.Type);
        }
    }
}