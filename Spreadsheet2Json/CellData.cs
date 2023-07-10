using System;
using System.Globalization;
using ClosedXML.Excel;

namespace Spreadsheet2Json
{
    /// <summary>
    /// Cell Data class.
    /// </summary>
    internal class CellData
    {
        private readonly AdditionalNumberFormat _additionalNumberFormat;
        private readonly string _cultureInfoName;
        private string _text = string.Empty;
        private XLDataType _type;

        /// <summary>
        /// Initializes a new instance of the <see cref="CellData"/> class.
        /// </summary>
        /// <param name="cultureInfoName">Culture Info Name.</param>
        public CellData(string cultureInfoName)
        {
            _additionalNumberFormat = new AdditionalNumberFormat(cultureInfoName);
            _cultureInfoName = cultureInfoName;
        }

        /// <summary>
        /// Sets cell.
        /// </summary>
        public IXLCell Cell
        {
            set { SetCell(value); }
        }

        /// <summary>
        /// Gets cell text.
        /// </summary>
        public string Text
        {
            get { return _text; }
        }

        /// <summary>
        /// Gets cell type.
        /// </summary>
        public XLDataType Type
        {
            get { return _type; }
        }

        private void SetCell(IXLCell cell)
        {
            // initialized value.
            _text = cell.GetFormattedString();
            _type = cell.Value.Type;

            if (_additionalNumberFormat.NumberFormatItemList.TryGetValue(cell.Style.NumberFormat.NumberFormatId, out NumberFormatItem numberFormatItem))
            {
                string format = numberFormatItem.Format;
                var nf = new ExcelNumberFormat.NumberFormat(format);
                _text = nf.Format(cell.Value.GetUnifiedNumber(), CultureInfo.InvariantCulture);

                _type = numberFormatItem.Type;
            }
            else
            {
                string format = cell.Style.NumberFormat.Format;
                if (format != string.Empty)
                {
                    if (_type == XLDataType.Number && format.Contains(@"[$-F400]"))
                    {
                        // Long time pattern
                        _text = DateTime.FromOADate(cell.Value.GetUnifiedNumber())
                            .ToString("T", CultureInfo.CreateSpecificCulture(_cultureInfoName));

                        _type = XLDataType.DateTime;
                    }

                    if (_type == XLDataType.Number && format.Contains(@"[$-409]"))
                    {
                        // Consider en-US date format
                        _type = XLDataType.DateTime;
                    }

                    if (_type == XLDataType.Number && format.Contains(@"[$-411]"))
                    {
                        // Consider ja-JP date format
                        _type = XLDataType.DateTime;

                        // TODO: Corresponds to the Japanese calendar
                    }

                    if (_type == XLDataType.Number && format.Contains(@"[$-F800]"))
                    {
                        // Long date pattern
                        _text = DateTime.FromOADate(cell.Value.GetUnifiedNumber())
                            .ToString("D", CultureInfo.CreateSpecificCulture(_cultureInfoName));

                        _type = XLDataType.DateTime;
                    }
                }
            }
        }
    }
}
