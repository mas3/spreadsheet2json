using System.CommandLine;
using System.CommandLine.Binding;

namespace Spreadsheet2Json
{
    /// <summary>
    /// Command Line Option swith class.
    /// </summary>
    internal class OptionSwitch
    {
        /// <summary>
        /// Gets or sets a value indicating whether option of include cell format.
        /// </summary>
        public bool IsIncludeCellFormat { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of include cell row and column.
        /// </summary>
        public bool IsIncludeCellRowColumn { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of include debug message.
        /// </summary>
        public bool IsDebugMode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of suppress encoding json.
        /// </summary>
        public bool IsNotEncoding { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of not include cell data.
        /// </summary>
        public bool IsNotIncludeCellData { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of not include workbook property data.
        /// </summary>
        public bool IsNotIncludeProperties { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of not include sheet information.
        /// </summary>
        public bool IsNotIncludeSheetInformation { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether option of suppress indenting json.
        /// </summary>
        public bool IsNotIndenting { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the represent sheet and cell data in object format.
        /// </summary>
        public bool IsObjectFormat { get; set; }
    }

    /// <summary>
    /// OptionSwitch binder class.
    /// </summary>
    internal class OptionSwitchBinder : BinderBase<OptionSwitch>
    {
        private readonly Option<bool> _isDebugMode;
        private readonly Option<bool> _isIncludeCellFormat;
        private readonly Option<bool> _isIncludeCellRowColumn;
        private readonly Option<bool> _isIncludeSheetInformation;
        private readonly Option<bool> _isNotEncoding;
        private readonly Option<bool> _isNotIncludeCellData;
        private readonly Option<bool> _isNotIncludeProperties;
        private readonly Option<bool> _isNotIndenting;
        private readonly Option<bool> _isObjectFormat;

        /// <summary>
        /// Initializes a new instance of the <see cref="OptionSwitchBinder"/> class.
        /// </summary>
        /// <param name="isDebugMode">Output debug message.</param>
        /// <param name="isIncludeCellFormat">Include cell format.</param>
        /// <param name="isIncludeCellRowColumn">Output cell row and column.</param>
        /// <param name="isIncludeSheetInformation">Output sheet information.</param>
        /// <param name="isNotEncoding">Suppress encoding json.</param>
        /// <param name="isNotIncludeCellData">Suppress output cell data.</param>
        /// <param name="isNotIncludeProperties">Suppress output workbook property data.</param>
        /// <param name="isNotIndent">Suppress indent json.</param>
        /// <param name="isObjectFormat">Represent sheet and cell data in object format.</param>
        public OptionSwitchBinder(Option<bool> isDebugMode, Option<bool> isIncludeCellFormat, Option<bool> isIncludeCellRowColumn, Option<bool> isIncludeSheetInformation, Option<bool> isNotEncoding, Option<bool> isNotIncludeCellData, Option<bool> isNotIncludeProperties, Option<bool> isNotIndent, Option<bool> isObjectFormat)
        {
            _isDebugMode = isDebugMode;
            _isIncludeCellFormat = isIncludeCellFormat;
            _isIncludeCellRowColumn = isIncludeCellRowColumn;
            _isIncludeSheetInformation = isIncludeSheetInformation;
            _isNotEncoding = isNotEncoding;
            _isNotIncludeCellData = isNotIncludeCellData;
            _isNotIncludeProperties = isNotIncludeProperties;
            _isNotIndenting = isNotIndent;
            _isObjectFormat = isObjectFormat;
        }

        /// <summary>
        /// Bind option value.
        /// </summary>
        /// <param name="bindingContext">binding context.</param>
        /// <returns>OptonSwitch instance.</returns>
        protected override OptionSwitch GetBoundValue(BindingContext bindingContext)
        {
            return new OptionSwitch
            {
                IsIncludeCellFormat = bindingContext.ParseResult.GetValueForOption(_isIncludeCellFormat),
                IsIncludeCellRowColumn = bindingContext.ParseResult.GetValueForOption(_isIncludeCellRowColumn),
                IsDebugMode = bindingContext.ParseResult.GetValueForOption(_isDebugMode),
                IsNotIncludeCellData = bindingContext.ParseResult.GetValueForOption(_isNotIncludeCellData),
                IsNotEncoding = bindingContext.ParseResult.GetValueForOption(_isNotEncoding),
                IsNotIndenting = bindingContext.ParseResult.GetValueForOption(_isNotIndenting),
                IsNotIncludeProperties = bindingContext.ParseResult.GetValueForOption(_isNotIncludeProperties),
                IsNotIncludeSheetInformation = bindingContext.ParseResult.GetValueForOption(_isIncludeSheetInformation),
                IsObjectFormat = bindingContext.ParseResult.GetValueForOption(_isObjectFormat),
            };
        }
    }
}
