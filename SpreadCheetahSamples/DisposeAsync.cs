using SpreadCheetah;

namespace SpreadCheetahSamples;

public static class DisposeAsync
{
    public static async Task Sample()
    {
        // From C# 8.0 and later, streams implement DisposeAsync and can be disposed with `await using`.
        await using var stream = File.Create("dispose-async.xlsx");

        // SpreadCheetah also similarly implements DisposeAsync.
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        Cell[] row =
        [
            new("Answer to the ultimate question:"),
            new(42)
        ];

        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.AddRowAsync(row);
        await spreadsheet.FinishAsync();
    }
}
