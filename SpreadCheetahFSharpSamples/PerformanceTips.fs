module SpreadCheetahFSharpSamples.PerformanceTips

open System.IO
open SpreadCheetah
open SpreadCheetah.Styling

let private getStudents() = [|
    ("Jakob", 22, "C")
    ("Emma", 20, "B")
    ("William", 23, "A")
    ("Sara", 22, "A")
    ("Lucas", 21, "D")
|]

let sample() = task {
    use stream = File.Create("performance-tips.xlsx")
    use! spreadsheet = Spreadsheet.CreateNewAsync(stream)

    do! spreadsheet.StartWorksheetAsync("Sheet 1")

    let headerStyle = Style()
    headerStyle.Font.Bold <- true
    let headerStyleId = spreadsheet.AddStyle(headerStyle)

    // `Cell` is the general purpose type for creating cells.
    // `StyledCell` can perform better than `Cell` for rows that only contain a value with optional styling.
    let headerRow = [|
        StyledCell("Student", headerStyleId)
        StyledCell("Age", headerStyleId)
        StyledCell("Grade", headerStyleId)
    |]

    // `DataCell` can perform even better for rows that only contain a value but with no styling.
    let dataRow = Array.zeroCreate<DataCell> headerRow.Length

    // A row can not contain a mixture of cell types, they must all either be a `Cell`, a `StyledCell`, or a `DataCell`.
    do! spreadsheet.AddRowAsync(headerRow)

    for name, age, grade in getStudents() do
        // Reusing an array or list can also avoid some memory allocations.
        dataRow[0] <- DataCell(name)
        dataRow[1] <- DataCell(age)
        dataRow[2] <- DataCell(grade)
        do! spreadsheet.AddRowAsync(dataRow)
}
