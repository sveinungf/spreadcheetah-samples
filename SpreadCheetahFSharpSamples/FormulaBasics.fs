module SpreadCheetahFSharpSamples.FormulaBasics

open System.IO
open SpreadCheetah

let sample() = task {
    use stream = File.Create("formula-basics.xlsx")
    use! spreadsheet = Spreadsheet.CreateNewAsync(stream)

    do! spreadsheet.StartWorksheetAsync("Sheet 1")

    // Add numeric values to cells A1, A2, and A3.
    do! spreadsheet.AddRowAsync([| Cell(10) |])
    do! spreadsheet.AddRowAsync([| Cell(20) |])
    do! spreadsheet.AddRowAsync([| Cell(30) |])

    // Example of using the SUM formula to add the values above.
    // NOTE: Formulas should NOT start with the `=` sign.
    let formulaA4 = Formula("SUM(A1:A3)")
    let cellA4 = Cell(formulaA4)
    do! spreadsheet.AddRowAsync([| cellA4 |])

    // Formula cells can optionally contain a cached value.
    // Formula cells without a cached value will otherwise be calculated when displayed in Excel.
    let formulaA5 = new Formula("AVERAGE(A1:A3)")
    let cachedValue = 20
    let cellA5 = Cell(formulaA5, cachedValue)
    do! spreadsheet.AddRowAsync([| cellA5 |])

    do! spreadsheet.FinishAsync()
}
