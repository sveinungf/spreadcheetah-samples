using SpreadCheetah;

namespace SpreadCheetahSamples;

public static class FormulaBasics
{
    public static async Task Sample()
    {
        using (var stream = File.Create("formula-basics.xlsx"))
        using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet 1");

            // Add numeric values to cells A1, A2, and A3.
            await spreadsheet.AddRowAsync(new[] { new Cell(10) });
            await spreadsheet.AddRowAsync(new[] { new Cell(20) });
            await spreadsheet.AddRowAsync(new[] { new Cell(30) });

            // Example of using the SUM formula to add the values above.
            // NOTE: Formulas should NOT start with the `=` sign.
            var formulaA4 = new Formula("SUM(A1:A3)");
            var cellA4 = new Cell(formulaA4);
            await spreadsheet.AddRowAsync(new[] { cellA4 });

            // Formula cells can optionally contain a cached value.
            // Formula cells without a cached value will otherwise be calculated when displayed in Excel.
            var formulaA5 = new Formula("AVERAGE(A1:A3)");
            var cachedValue = 20;
            var cellA5 = new Cell(formulaA5, cachedValue);
            await spreadsheet.AddRowAsync(new[] { cellA5 });

            await spreadsheet.FinishAsync();
        }
    }
}
