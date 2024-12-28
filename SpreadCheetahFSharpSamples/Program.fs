module SpreadCheetahFSharpSamples.Program

let main() = task {
    do! DataValidations.sample()
    do! DateTimeAndFormatting.sample()
    do! DisposeAsync.sample()
    do! FormulaBasics.sample()
    do! PerformanceTips.sample()
    do! StylingBasics.sample()
    do! WriteToFile.sample()
}

main().GetAwaiter().GetResult()
printfn "Done!"
