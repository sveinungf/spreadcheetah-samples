using SpreadCheetah;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace SpreadCheetahSamples
{
    public static class StylingBasics
    {
        public static async Task Sample()
        {
            using (var stream = File.Create("styling-basics.xlsx"))
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                // Optionally set column widths.
                var worksheetOptions = new WorksheetOptions();
                worksheetOptions.Column(1).Width = 100;
                worksheetOptions.Column(2).Width = 80;

                await spreadsheet.StartWorksheetAsync("Sheet 1", worksheetOptions);

                // Defining a style with a custom font.
                // Style properties that have not been set will get the values that are default in Excel.
                var questionStyle = new Style();
                questionStyle.Font.Bold = true;
                questionStyle.Font.Size = 20;

                // Defining a style with font and fill color.
                // Colors are specified using System.Drawing.Color.
                var answerStyle = new Style();
                answerStyle.Fill.Color = Color.Green;
                answerStyle.Font.Color = Color.FromArgb(100, 150, 200);

                // We need style IDs to use the styles. Use `AddStyle` to get a style ID.
                var questionStyleId = spreadsheet.AddStyle(questionStyle);
                var answerStyleId = spreadsheet.AddStyle(answerStyle);

                // Pass the style ID when creating the cells.
                var row1 = new[]
                {
                    new Cell("Highest mountain?", questionStyleId),
                    new Cell("Mount Everest", answerStyleId)
                };

                // Existing style IDs can be reused across cells.
                var row2 = new[]
                {
                    new Cell("Longest river?", questionStyleId),
                    new Cell("The Nile", answerStyleId)
                };

                await spreadsheet.AddRowAsync(row1);
                await spreadsheet.AddRowAsync(row2);

                await spreadsheet.FinishAsync();
            }
        }
    }
}
