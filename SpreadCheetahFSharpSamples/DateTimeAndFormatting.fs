module SpreadCheetahFSharpSamples.DateTimeAndFormatting

open System
open System.IO
open SpreadCheetah
open SpreadCheetah.Styling
open SpreadCheetah.Worksheets

let sample() = task {
    use stream = File.Create("datetime-and-formatting.xlsx")
    use! spreadsheet = Spreadsheet.CreateNewAsync(stream)

    let options = WorksheetOptions()
    options.Column(1).Width <- 20
    do! spreadsheet.StartWorksheetAsync("Sheet 1", options)

    let dateTime = DateTime(2022, 10, 18, 11, 26, 34)

    // Example of writing a DateTime with a custom number format. The date will be displayed as "18.10.2022".
    // Note that the format must be an Excel format code. More information here: https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
    let style1 = Style(Format = NumberFormat.Custom("DD.MM.YYYY"))
    let style1Id = spreadsheet.AddStyle(style1)
    let cellA1 = Cell(dateTime, style1Id)
    do! spreadsheet.AddRowAsync([| cellA1 |])

    // Note that some characters have special meaning in the format codes. An example is the 'h' character, which signifies the hour.
    // Text can be escaped by enclosing them in double quotation marks. Here is an example of displaying the date as "18th":
    let style2 = Style(Format = NumberFormat.Custom("D\"th\""))
    let style2Id = spreadsheet.AddStyle(style2)
    let cellA2 = Cell(dateTime, style2Id)
    do! spreadsheet.AddRowAsync([| cellA2 |])

    // Also note that how some parts are displayed can depend on the regional/language setting of Excel.
    // This example will be shown as "October" in Excel when English (US) is the chosen language.
    // It will be shown as "oktober" in Excel when Norwegian is the chosen language.
    let style3 = Style(Format = NumberFormat.Custom("MMMM"))
    let style3Id = spreadsheet.AddStyle(style3)
    let cellA3 = Cell(dateTime, style3Id)
    do! spreadsheet.AddRowAsync([| cellA3 |])

    // When no style or number format has been specified, the DateTime will by default be displayed as "2022-10-18 11:26:34".
    // The default can be overriden by setting DefaultDateTimeNumberFormat on SpreadCheetahOptions when creating the spreadsheet.
    let cellA4 = Cell(dateTime)
    do! spreadsheet.AddRowAsync([| cellA4 |])

    do! spreadsheet.FinishAsync()
}
