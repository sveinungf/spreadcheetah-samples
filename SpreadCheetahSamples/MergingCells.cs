using SpreadCheetah;
using SpreadCheetah.Styling;

namespace SpreadCheetahSamples;

public static class MergingCells
{
    public static async Task Sample()
    {
        await using var stream = File.Create("merging-cells.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");

        var centerAlignStyle = new Style
        {
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignment.Center,
                Vertical = VerticalAlignment.Center
            }
        };

        var centerAlignStyleId = spreadsheet.AddStyle(centerAlignStyle);

        // Row 1
        await spreadsheet.AddRowAsync([new Cell("Report title")]);
        spreadsheet.MergeCells("A1:C1");

        // Row 2
        await spreadsheet.AddRowAsync([]);

        // Row 3
        await spreadsheet.AddRowAsync([new Cell("Content...", centerAlignStyleId)]);
        spreadsheet.MergeCells("A3:D6");

        await spreadsheet.FinishAsync();
    }
}
