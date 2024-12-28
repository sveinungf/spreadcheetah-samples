module SpreadCheetahFSharpSamples.WriteToFile

open System.IO
open SpreadCheetah

let sample() = task {
    // SpreadCheetah can write to any writeable stream.
    // To write to a file, start by creating a file stream.
    use stream = File.Create("write-to-file.xlsx")
    use! spreadsheet = Spreadsheet.CreateNewAsync(stream)

    // A spreadsheet must contain at least one worksheet.
    do! spreadsheet.StartWorksheetAsync("Sheet 1")

    // Cells are inserted row by row.
    let row = ResizeArray()
    row.Add(Cell("Answer to the ultimate question:"))
    row.Add(Cell(42))

    // Rows are inserted from top to bottom.
    do! spreadsheet.AddRowAsync(row)

    // Remember to call Finish before disposing.
    // This is important to properly finalize the XLSX file.
    do! spreadsheet.FinishAsync()
}
