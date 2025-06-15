using SpreadCheetah;
using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetahSamples;

public static class FormulaBasics
{
    public static async Task Sample()
    {
        await using var stream = File.Create("formula-basics.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        // Add numeric values to cells A1, A2, and A3.
        await spreadsheet.AddRowAsync([new Cell(10)]);
        await spreadsheet.AddRowAsync([new Cell(20)]);
        await spreadsheet.AddRowAsync([new Cell(30)]);

        // Example of using the SUM formula to add the values above.
        // NOTE: Formulas should NOT start with the `=` sign.
        var formulaA4 = new Formula("SUM(A1:A3)");
        var cellA4 = new Cell(formulaA4);
        await spreadsheet.AddRowAsync([cellA4]);

        // Formula cells can optionally contain a cached value.
        // Formula cells without a cached value will otherwise be calculated when displayed in Excel.
        var formulaA5 = new Formula("AVERAGE(A1:A3)");
        var cachedValue = 20;
        var cellA5 = new Cell(formulaA5, cachedValue);
        await spreadsheet.AddRowAsync([cellA5]);

        // Creating hyperlink formulas can be done using one of the Hyperlink helper methods.
        var uri = new Uri("https://github.com/sveinungf/spreadcheetah");
        var formulaA6 = Formula.Hyperlink(uri, "GitHub Repository");

        // You can use a style to make it look like a clickable link.
        var hyperlinkStyle = new Style
        {
            Font = new Font
            {
                Color = Color.Blue,
                Underline = Underline.Single
            }
        };
        var hyperlinkStyleId = spreadsheet.AddStyle(hyperlinkStyle);

        var cellA6 = new Cell(formulaA6, hyperlinkStyleId);
        await spreadsheet.AddRowAsync([cellA6]);

        await spreadsheet.FinishAsync();
    }
}
