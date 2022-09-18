using SpreadCheetah;
using SpreadCheetah.Styling;

namespace SpreadCheetahSamples;

public static class DateTimeAndFormatting
{
    public static async Task Sample()
    {
        await using var stream = File.Create("datetime-and-formatting.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        var dateTime = new DateTime(2022, 10, 18, 11, 26, 34);

        // Example of writing a DateTime with a custom number format. The date will be displayed as "18.10.2022".
        // Note that the format must be an Excel format code. More information here: https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        var style1 = new Style { NumberFormat = "DD.MM.YYYY" };
        var style1Id = spreadsheet.AddStyle(style1);
        var cellA1 = new Cell(dateTime, style1Id);
        await spreadsheet.AddRowAsync(new[] { cellA1 });

        // Note that some characters have special meaning in the format codes. An example is the 'h' character, which signifies the hour.
        // Text can be escaped by enclosing them in double quotation marks. Here is an example of displaying the date as "18th":
        var style2 = new Style { NumberFormat = "D\"th\"" };
        var style2Id = spreadsheet.AddStyle(style2);
        var cellA2 = new Cell(dateTime, style2Id);
        await spreadsheet.AddRowAsync(new[] { cellA2 });

        // Also note that how some parts are displayed can depend on the regional/language setting of Excel.
        // This example will be shown as "October" in Excel when English (US) is the chosen language.
        // It will be shown as "oktober" in Excel when Norwegian is the chosen language.
        var style3 = new Style { NumberFormat = "MMMM" };
        var style3Id = spreadsheet.AddStyle(style3);
        var cellA3 = new Cell(dateTime, style3Id);
        await spreadsheet.AddRowAsync(new[] { cellA3 });

        // When no style or number format has been specified, the DateTime will by default be displayed as "2022-10-18 11:26:34".
        // The default can be overriden by setting DefaultDateTimeNumberFormat on SpreadCheetahOptions when creating the spreadsheet.
        var cellA4 = new Cell(dateTime);
        await spreadsheet.AddRowAsync(new[] { cellA4 });

        await spreadsheet.FinishAsync();
    }
}
