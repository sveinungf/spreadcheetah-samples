using SpreadCheetah;
using SpreadCheetah.ConditionalFormatting;
using System.Drawing;

namespace SpreadCheetahSamples.ConditionalFormatting;

public static class ConditionalFormatUniqueValues
{
    public static async Task Sample()
    {
        await using var outputStream = File.Create("conditional-format-unique-values.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new ConditionalFormatStyle { Fill = { Color = Color.Yellow } };
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);

        const string cellRange = "A1:E1";
        spreadsheet.AddConditionalFormatRule(cellRange, rule);

        DataCell[] row = [new(1), new(2), new(3), new(3), new(1)];
        await spreadsheet.AddRowAsync(row);

        await spreadsheet.FinishAsync();
    }
}
