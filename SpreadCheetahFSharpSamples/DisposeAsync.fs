module SpreadCheetahFSharpSamples.DisposeAsync

open System.IO
open SpreadCheetah

let sample() = task {
    use stream = File.Create("dispose-async.xlsx")

    // SpreadCheetah also similarly implements DisposeAsync.
    use! spreadsheet = Spreadsheet.CreateNewAsync(stream)

    let row = [|
        Cell("Answer to the ultimate question:")
        Cell(42)
    |]

    do! spreadsheet.StartWorksheetAsync("Sheet 1")
    do! spreadsheet.AddRowAsync(row)
    do! spreadsheet.FinishAsync()
}
