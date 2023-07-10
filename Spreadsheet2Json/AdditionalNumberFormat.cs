using System.Collections.Generic;
using ClosedXML.Excel;

namespace Spreadsheet2Json
{
    /// <summary>
    /// Additional Number Format class.
    /// </summary>
    internal class AdditionalNumberFormat
    {
        private readonly Dictionary<int, NumberFormatItem> _numberFormatItemList;

        private readonly Dictionary<string, Dictionary<int, NumberFormatItem>> numberFormatItemListByCountry = new()
        {
            // cf. https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
            ["ja-JP"] =
                new Dictionary<int, NumberFormatItem>()
                {
                    { 14, new NumberFormatItem(@"yyyy/m/d", XLDataType.DateTime) },
                    { 30, new NumberFormatItem(@"m/d/yy", XLDataType.DateTime) },
                    { 31, new NumberFormatItem(@"yyyy""年""m""月""d""日""", XLDataType.DateTime) },
                    { 55, new NumberFormatItem(@"yyyy""年""m""月""", XLDataType.DateTime) },
                    { 56, new NumberFormatItem(@"m""月""d""日""", XLDataType.DateTime) },
                },
        };

        /// <summary>
        /// Initializes a new instance of the <see cref="AdditionalNumberFormat"/> class.
        /// </summary>
        /// <param name="cultureInfoName">Culture Info Name.</param>
        public AdditionalNumberFormat(string cultureInfoName)
        {
            if (numberFormatItemListByCountry.ContainsKey(cultureInfoName))
            {
                _numberFormatItemList = numberFormatItemListByCountry[cultureInfoName];
            }
            else
            {
                _numberFormatItemList = new Dictionary<int, NumberFormatItem>();
            }
        }

        /// <summary>
        /// Gets Number format item list.
        /// </summary>
        public Dictionary<int, NumberFormatItem> NumberFormatItemList
        {
            get { return _numberFormatItemList; }
        }
    }
}
