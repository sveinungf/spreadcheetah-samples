using SpreadCheetah;
using SpreadCheetah.Styling;

namespace SpreadCheetahSamples;

public static class PerformanceTips
{
    public static async Task Sample()
    {
        await using var stream = File.Create("performance-tips.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        var headerStyle = new Style();
        headerStyle.Font.Bold = true;
        var headerStyleId = spreadsheet.AddStyle(headerStyle);

        // `Cell` is the general purpose type for creating cells.
        // `StyledCell` can perform better than `Cell` for rows that only contain a value with optional styling.
        var headerRow = new[]
        {
            new StyledCell("Student", headerStyleId),
            new StyledCell("Age", headerStyleId),
            new StyledCell("Grade", headerStyleId)
        };

        // `DataCell` can perform even better for rows that only contain a value but with no styling.
        var dataRow = new DataCell[headerRow.Length];

        // A row can not contain a mixture of cell types, they must all either be a `Cell`, a `StyledCell`, or a `DataCell`.
        await spreadsheet.AddRowAsync(headerRow);

        foreach (var (Name, Age, Grade) in GetStudents())
        {
            // Reusing an array or list can also avoid some memory allocations.
            dataRow[0] = new DataCell(Name);
            dataRow[1] = new DataCell(Age);
            dataRow[2] = new DataCell(Grade);
            await spreadsheet.AddRowAsync(dataRow);
        }

        await spreadsheet.FinishAsync();
    }

    private static (string Name, int Age, string Grade)[] GetStudents() =>
    [
        ("Jakob", 22, "C"),
        ("Emma", 20, "B"),
        ("William", 23, "A"),
        ("Sara", 22, "A"),
        ("Lucas", 21, "D")
    ];
}
