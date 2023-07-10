using System;
using System.CommandLine;
using System.Globalization;
using System.IO;
using Spreadsheet2Json;

// Main
{
    Option<bool> cellFormatAgument = new(
        name: "--cell-format",
        description: $"Include cell format data.",
        getDefaultValue: () => false);

    Option<bool> cellRowColumnAgument = new(
        name: "--cell-row-column",
        description: $"Include cell row and column address data.",
        getDefaultValue: () => false);

    Option<bool> debugAgument = new(
        name: "--debug",
        description: $"Show debug message.",
        getDefaultValue: () => false);

    Option<FileInfo?> inputFileOption = new(
        name: "--input",
        description: "Input xlsx filename.")
    { IsRequired = true, ArgumentHelpName = "filename" };
    inputFileOption.AddAlias("-i");

    Option<bool> cellDataAgument = new(
        name: "--no-cell-data",
        description: $"Suppress including cell data.",
        getDefaultValue: () => false);

    Option<bool> encodeAgument = new(
        name: "--no-encode",
        description: $"Suppress encoding process of JSON string.",
        getDefaultValue: () => false);

    Option<bool> indentAgument = new(
        name: "--no-indent",
        description: $"Suppress indent process of JSON string.",
        getDefaultValue: () => false);

    Option<bool> propertyAgument = new(
        name: "--no-properties",
        description: $"Suppress including workbook property data.",
        getDefaultValue: () => false);

    Option<bool> objectFormatAgument = new(
        name: "--object-format",
        description: $"Represent sheet and cell data in object format.",
        getDefaultValue: () => false);

    Option<FileInfo?> outpuFileOption = new(
        name: "--output",
        description: $"Output json filename. If omitted, output JSON string to standard output.")
    { ArgumentHelpName = "filename" };
    outpuFileOption.AddAlias("-o");

    Option<bool> sheetInformationAgument = new(
        name: "--sheet-info",
        description: $"Include sheet information.",
        getDefaultValue: () => false);

    Option<string> sheetsAgument = new(
        name: "--sheets",
        description: $"Target sheet names. Separate multiple sheets with colons. e.g., sheet1:sheet2");

    RootCommand rootCommand = new("Spreadsheet (*.xlsx file) to JSON command.");
    rootCommand.AddOption(inputFileOption);
    rootCommand.AddOption(outpuFileOption);
    rootCommand.AddOption(cellFormatAgument);
    rootCommand.AddOption(cellRowColumnAgument);
    rootCommand.AddOption(debugAgument);
    rootCommand.AddOption(cellDataAgument);
    rootCommand.AddOption(encodeAgument);
    rootCommand.AddOption(indentAgument);
    rootCommand.AddOption(propertyAgument);
    rootCommand.AddOption(objectFormatAgument);
    rootCommand.AddOption(sheetInformationAgument);
    rootCommand.AddOption(sheetsAgument);

    rootCommand.SetHandler(
        (inputFile, outputFile, sheetNames, optionSwitch) =>
        {
            Execute(inputFile, outputFile, sheetNames, optionSwitch);
        },
        inputFileOption,
        outpuFileOption,
        sheetsAgument,
        new OptionSwitchBinder(debugAgument, cellFormatAgument, cellRowColumnAgument, sheetInformationAgument, encodeAgument, cellDataAgument, propertyAgument, indentAgument, objectFormatAgument));

    return await rootCommand.InvokeAsync(args);
}

static void Execute(FileInfo? inputFile, FileInfo? outputFile, string sheetNames, OptionSwitch optionSwitch)
{
    SpreadsheetTranslator translator = new(CultureInfo.CurrentCulture.Name)
    {
        Encoded = !optionSwitch.IsNotEncoding,
        IsIncludeCellData = !optionSwitch.IsNotIncludeCellData,
        IsIncludeCellFormat = optionSwitch.IsIncludeCellFormat,
        IsIncludeCellRowColumn = optionSwitch.IsIncludeCellRowColumn,
        IsIncludeProperties = !optionSwitch.IsNotIncludeProperties,
        IsIncludeSheetInformation = optionSwitch.IsNotIncludeSheetInformation,
        IsObjectFormat = optionSwitch.IsObjectFormat,
        Indented = !optionSwitch.IsNotIndenting,
        InputFilePath = inputFile?.FullName ?? string.Empty,
        SheetNames = sheetNames ?? string.Empty,
    };

    try
    {
        string jsonString = translator.GetJsonString();
        if (outputFile != null)
        {
            System.Text.UTF8Encoding utf8WithoutBom = new(false);
            File.WriteAllText(outputFile.FullName, jsonString, utf8WithoutBom);
        }
        else
        {
            Console.WriteLine(jsonString);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
        if (optionSwitch.IsDebugMode)
        {
            Console.WriteLine(ex.StackTrace);
        }

        Environment.Exit(1);
    }
}
