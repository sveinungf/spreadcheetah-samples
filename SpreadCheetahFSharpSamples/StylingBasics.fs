module SpreadCheetahFSharpSamples.StylingBasics

open System.Drawing
open System.IO
open SpreadCheetah
open SpreadCheetah.Styling
open SpreadCheetah.Worksheets

let sample() = task {
    use stream = File.Create("styling-basics.xlsx")
    use! spreadsheet = Spreadsheet.CreateNewAsync(stream)

    // Optionally set column widths.
    let worksheetOptions = WorksheetOptions()
    worksheetOptions.Column(1).Width <- 100
    worksheetOptions.Column(2).Width <- 80

    do! spreadsheet.StartWorksheetAsync("Sheet 1", worksheetOptions)

    // Defining a style with a custom font.
    // Style properties that have not been set will get the values that are default in Excel.
    let questionStyle = Style()
    questionStyle.Font.Bold <- true
    questionStyle.Font.Size <- 20

    // Defining a style with font and fill color.
    // Colors are specified using System.Drawing.Color.
    let answerStyle = Style()
    answerStyle.Fill.Color <- Color.Green
    answerStyle.Font.Color <- Color.FromArgb(100, 150, 200)

    // We need style IDs to use the styles. Use `AddStyle` to get a style ID.
    let questionStyleId = spreadsheet.AddStyle(questionStyle)
    let answerStyleId = spreadsheet.AddStyle(answerStyle)

    // Pass the style ID when creating the cells.
    let row1 = [|
        Cell("Highest mountain?", questionStyleId)
        Cell("Mount Everest", answerStyleId)
    |]

    // Existing style IDs can be reused across cells.
    let row2 = [|
        Cell("Longest river?", questionStyleId)
        Cell("The Nile", answerStyleId)
    |]

    // Optionally set row height.
    let rowOptions = RowOptions(Height = 25)

    do! spreadsheet.AddRowAsync(row1, rowOptions)
    do! spreadsheet.AddRowAsync(row2, rowOptions)

    do! spreadsheet.FinishAsync()
}
