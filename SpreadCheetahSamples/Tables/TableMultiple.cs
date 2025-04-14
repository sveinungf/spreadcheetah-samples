using SpreadCheetah;
using SpreadCheetah.Tables;
using SpreadCheetah.Worksheets;

namespace SpreadCheetahSamples.Tables;

public static class TableMultiple
{
    public static async Task Sample()
    {
        await using var outputStream = File.Create("table-multiple.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        var worksheetOptions = new WorksheetOptions();
        worksheetOptions.Column(2).Width = 18;
        await spreadsheet.StartWorksheetAsync("Sheet", worksheetOptions);

        string[] headerNames = ["City", "Temperature (°C)"];

        var table1 = new Table(TableStyle.Dark5);
        spreadsheet.StartTable(table1);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync([new DataCell("Paris"), new DataCell(14)]);
        await spreadsheet.AddRowAsync([new DataCell("Bangkok"), new DataCell(34)]);
        await spreadsheet.FinishTableAsync();

        await spreadsheet.AddRowAsync([]);

        var table2 = new Table(TableStyle.Dark6);
        spreadsheet.StartTable(table2);
        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync([new DataCell("Dakar"), new DataCell(21)]);
        await spreadsheet.AddRowAsync([new DataCell("Lima"), new DataCell(24)]);
        await spreadsheet.FinishTableAsync();

        await spreadsheet.FinishAsync();
    }
}
