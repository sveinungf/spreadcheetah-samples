using SpreadCheetah;
using SpreadCheetah.Styling;
using System.IO;
using System.Threading.Tasks;

namespace SpreadCheetahSamples
{
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

            // `StyledCell` can perform better than `Cell` for rows that only contains a value with styling.
            var headerRow = new[]
            {
                new StyledCell("Student", headerStyleId),
                new StyledCell("Age", headerStyleId),
                new StyledCell("Grade", headerStyleId)
            };

            // `DataCell` can perform even better for rows that only contains a value with no styling.
            // If all rows have the same number of columns, reusing an array/list can also avoid some memory allocations.
            var row = new DataCell[headerRow.Length];

            foreach (var (Name, Age, Grade) in GetStudents())
            {
                row[0] = new DataCell(Name);
                row[1] = new DataCell(Age);
                row[2] = new DataCell(Grade);
                await spreadsheet.AddRowAsync(row);
            }

            await spreadsheet.FinishAsync();
        }

        private static (string Name, int Age, string Grade)[] GetStudents() => new[]
        {
            ("Jakob", 22, "C"),
            ("Emma", 20, "B"),
            ("William", 23, "A"),
            ("Sara", 22, "A"),
            ("Lucas", 21, "D")
        };
    }
}
