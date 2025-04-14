using SpreadCheetah;
using SpreadCheetah.Tables;
using SpreadCheetah.Worksheets;

namespace SpreadCheetahSamples.Tables;

public static class TableTotalRow
{
    public static async Task Sample()
    {
        await using var outputStream = File.Create("table-total-row.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        var worksheetOptions = new WorksheetOptions();
        worksheetOptions.Column(2).Width = 18;
        await spreadsheet.StartWorksheetAsync("Sheet", worksheetOptions);

        var table = new Table(TableStyle.Medium3);
        table.Column(1).TotalRowLabel = "Average";
        table.Column(2).TotalRowFunction = TableTotalRowFunction.Average;
        spreadsheet.StartTable(table);

        string[] headerNames = ["City", "Temperature (°C)"];
        DataCell[] paris = [new DataCell("Paris"), new DataCell(14)];
        DataCell[] bangkok = [new DataCell("Bangkok"), new DataCell(34)];

        await spreadsheet.AddHeaderRowAsync(headerNames);
        await spreadsheet.AddRowAsync(paris);
        await spreadsheet.AddRowAsync(bangkok);

        await spreadsheet.FinishAsync();
    }
}
