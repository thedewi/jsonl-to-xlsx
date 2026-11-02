using ClosedXML.Excel;
using System.CommandLine;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

Argument<FileInfo> outFileArgument = new("out-file")
{
    Description = "Path of the Excel workbook to write",
    Validators =
    {
        result =>
        {
            if (result.Tokens[0].Value is ['-', ..] option)
                result.AddError($"Unrecognized option '{option}'.");
        }
    }
};
RootCommand rootCommand = new("Convert JSON Lines table to Excel .xlsx")
{
    Arguments = { outFileArgument }
};

rootCommand.SetAction(parseResult =>
{
    var file = parseResult.GetRequiredValue(outFileArgument);
    if (file.Extension.Length == 0)
        file = new FileInfo(file.FullName + ".xlsx");
    if (file.Extension.ToLowerInvariant() is not (".xlsx" or ".xlsm" or ".xltx" or ".xltm"))
        throw new ArgumentException($"Extension '{file.Extension}' is not a valid workbook extension.");

    using XLWorkbook workbook = new();
    var sheetName = SheetNameUnWhitelist().Replace(Path.GetFileNameWithoutExtension(file.Name), "")
        .Trim() is { Length: > 0 } stripped
        ? stripped[..Math.Min(stripped.Length, 31)]
        : "Sheet1";
    var worksheet = workbook.AddWorksheet(sheetName);

    worksheet.Row(1).Style.Font.Bold = true;
    worksheet.SheetView.FreezeRows(1);

    var propertyOffsets = new Dictionary<string, int>();
    int lineNum = 0;
    Console.Error.WriteLine("Reading from stdin.");
    while (Console.In.ReadLine() is { } line)
    {
        ++lineNum;
        using var record = JsonDocument.Parse(line);
        foreach (var property in record.RootElement.EnumerateObject())
        {
            if (!propertyOffsets.TryGetValue(property.Name, out var offset))
            {
                offset = propertyOffsets[property.Name] = propertyOffsets.Count;
                worksheet.Cell(1, offset + 1).SetValue(property.Name);
            }

            XLCellValue value = property.Value.ValueKind switch
            {
                JsonValueKind.Object or JsonValueKind.Array => property.Value.GetRawText(),
                JsonValueKind.String => property.Value.GetString(),
                JsonValueKind.Number => property.Value.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null or JsonValueKind.Undefined => Blank.Value,
                _ => throw new InvalidOperationException($"Unexpected JsonValueKind in input line: {line}")
            };

            worksheet.Cell(lineNum + 1, offset + 1).SetValue(value);
        }

        if (lineNum % 1000 == 0)
            Console.Error.WriteLine($"Progress: converted {lineNum} lines.");
    }

    if (propertyOffsets.Count == 0)
        throw new InvalidOperationException($"No JSON properties found in input.");

    Console.Error.WriteLine($"Finished converting {lineNum} lines. Formatting and saving.");

    worksheet.RangeUsed()!.SetAutoFilter();
    worksheet.ColumnsUsed().AdjustToContents();
    workbook.SaveAs(file.FullName);

    // Reopen document to hide unused areas
    using var doc = SpreadsheetDocument.Open(file.FullName, true);
    if (doc.WorkbookPart is { } docBookPart)
    {
        var docSheet = docBookPart.WorksheetParts.First().Worksheet;
        docSheet.SheetFormatProperties!.ZeroHeight = true;
        if (docSheet.GetFirstChild<Columns>() is { } docCols)
            docCols.Append(new Column { Min = (uint)propertyOffsets.Count + 1, Max = 16384, Hidden = true });
    }
});

return await rootCommand.Parse(args).InvokeAsync();

internal static partial class Program
{
    [GeneratedRegex(@"[^\p{L}\p{N} _.()-]")]
    private static partial Regex SheetNameUnWhitelist();
}