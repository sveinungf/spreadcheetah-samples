using SpreadCheetah;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace SpreadCheetahSamples
{
    public static class WriteToFile
    {
        public static async Task Sample()
        {
            // SpreadCheetah can write to any writeable stream.
            // To write to a file, start by creating a file stream.
            using (var stream = File.Create("write-to-file.xlsx"))
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                // A spreadsheet must contain at least one worksheet.
                await spreadsheet.StartWorksheetAsync("Sheet 1");

                // Cells are inserted row by row.
                var row = new List<Cell>();
                row.Add(new Cell("Answer to the ultimate question:"));
                row.Add(new Cell(42));

                // Rows are inserted from top to bottom.
                await spreadsheet.AddRowAsync(row);

                // Remember to call Finish before disposing.
                // This is important to properly finalize the XLSX file.
                await spreadsheet.FinishAsync();
            }
        }
    }
}
