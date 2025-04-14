using SpreadCheetah;
using SpreadCheetah.Tables;
using SpreadCheetah.Worksheets;

namespace SpreadCheetahSamples.Tables;

public static class TableCalculatedColumn
{
    public static async Task Sample()
    {
        await using var outputStream = File.Create("table-calculated-column.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        var worksheetOptions = new WorksheetOptions();
        worksheetOptions.Column(1).Width = 18;
        await spreadsheet.StartWorksheetAsync("Sheet", worksheetOptions);

        var table = new Table(TableStyle.Medium4, "MyProductTable");
        spreadsheet.StartTable(table);

        string[] headerNames = ["Product", "Qtr 1", "Qtr 2", "Total"];
        var grandTotalFormula = new Formula("SUM(MyProductTable[[#This Row],[Qtr 1]:[Qtr 2]])");
        Cell[] chocolate = [new("Chocolate"), new(744), new(162), new Cell(grandTotalFormula)];
        Cell[] tomatoes = [new("Tomatoes"), new(345), new(377), new Cell(grandTotalFormula)];

        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync(chocolate);
        await spreadsheet.AddRowAsync(tomatoes);

        await spreadsheet.FinishAsync();
    }
}
