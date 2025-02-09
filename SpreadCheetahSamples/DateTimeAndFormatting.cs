using SpreadCheetah;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetahSamples;

public static class DateTimeAndFormatting
{
    public static async Task Sample()
    {
        await using var stream = File.Create("datetime-and-formatting.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        var options = new WorksheetOptions();
        options.Column(1).Width = 20;
        await spreadsheet.StartWorksheetAsync("Sheet 1", options);

        var dateTime = new DateTime(2022, 10, 18, 11, 26, 34);

        // Example of writing a DateTime with a custom number format. The date will be displayed as "18.10.2022".
        // Note that the format must be an Excel format code. More information here: https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        var style1 = new Style { Format = NumberFormat.Custom("DD.MM.YYYY") };
        var style1Id = spreadsheet.AddStyle(style1);
        var cellA1 = new Cell(dateTime, style1Id);
        await spreadsheet.AddRowAsync([cellA1]);

        // Note that some characters have special meaning in the format codes. An example is the 'h' character, which signifies the hour.
        // Text can be escaped by enclosing them in double quotation marks. Here is an example of displaying the date as "18th":
        var style2 = new Style { Format = NumberFormat.Custom("D\"th\"") };
        var style2Id = spreadsheet.AddStyle(style2);
        var cellA2 = new Cell(dateTime, style2Id);
        await spreadsheet.AddRowAsync([cellA2]);

        // Also note that how some parts are displayed can depend on the regional/language setting of Excel.
        // This example will be shown as "October" in Excel when English (US) is the chosen language.
        // It will be shown as "oktober" in Excel when Norwegian is the chosen language.
        var style3 = new Style { Format = NumberFormat.Custom("MMMM") };
        var style3Id = spreadsheet.AddStyle(style3);
        var cellA3 = new Cell(dateTime, style3Id);
        await spreadsheet.AddRowAsync([cellA3]);

        // When no style or number format has been specified, the DateTime will by default be displayed as "2022-10-18 11:26:34".
        // The default can be overriden by setting DefaultDateTimeNumberFormat on SpreadCheetahOptions when creating the spreadsheet.
        var cellA4 = new Cell(dateTime);
        await spreadsheet.AddRowAsync([cellA4]);

        // Example of conditional coloring - values that are less or equal to 100 will be red, and all others (greater than 100) will be blue
        // Note: Only the following colors can be used: [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]
        var style4 = new Style { Format = NumberFormat.Custom("[Red][<=100];[Blue][>100]") };
        var style4Id = spreadsheet.AddStyle(style4);
        var cellA5 = new Cell(99, style4Id);
        await spreadsheet.AddRowAsync(new[] { cellA5 });
        var cellA6 = new Cell(101, style4Id);
        await spreadsheet.AddRowAsync(new[] { cellA6 });

        await spreadsheet.FinishAsync();
    }
}
