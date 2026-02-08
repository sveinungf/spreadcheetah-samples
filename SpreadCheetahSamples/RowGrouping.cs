using SpreadCheetah;
using SpreadCheetah.Worksheets;

namespace SpreadCheetahSamples;

public static class RowGrouping
{
    public static async Task Sample()
    {
        await using var stream = File.Create("row-grouping.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        // While not strictly required, it's recommended to set the MaxRowOutlineLevel when using row grouping.
        // This allows Excel to know how many levels of grouping to expect and can help with performance.
        var worksheetOptions = new WorksheetOptions { MaxRowOutlineLevel = 2 };

        await spreadsheet.StartWorksheetAsync("Sheet 1", worksheetOptions);

        var group1Options = new RowOptions { OutlineLevel = 1 };
        var group2Options = new RowOptions { OutlineLevel = 2 };

        await spreadsheet.AddRowAsync([new Cell(10)]);
        await spreadsheet.AddRowAsync([new Cell(20)], group1Options);
        await spreadsheet.AddRowAsync([new Cell(30)], group2Options);
        await spreadsheet.AddRowAsync([new Cell(40)], group2Options);
        await spreadsheet.AddRowAsync([new Cell(50)]);

        await spreadsheet.FinishAsync();
    }
}
