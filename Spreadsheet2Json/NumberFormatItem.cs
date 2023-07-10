using ClosedXML.Excel;

namespace Spreadsheet2Json
{
    /// <summary>
    /// Struct of Number format item.
    /// </summary>
    internal struct NumberFormatItem
    {
        /// <summary>
        /// Format string.
        /// </summary>
        public string Format;

        /// <summary>
        /// Cell type.
        /// </summary>
        public XLDataType Type;

        /// <summary>
        /// Initializes a new instance of the <see cref="NumberFormatItem"/> struct.
        /// </summary>
        /// <param name="format">Number format.</param>
        /// <param name="type">Cell type.</param>
        public NumberFormatItem(string format, XLDataType type)
        {
            Format = format;
            Type = type;
        }
    }
}
