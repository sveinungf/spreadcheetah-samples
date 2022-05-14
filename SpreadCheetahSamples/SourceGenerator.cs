using SpreadCheetah;
using SpreadCheetah.SourceGeneration;

namespace SpreadCheetahSamples;

// A plain old C# class which we want to add as a row in a worksheet.
// The source generator will pick the properties with public getters.
// The order of the properties will decide the order of the cell values.
public class Person
{
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public int Age { get; set; }
}


// The source generator is used in a similar way as the System.Text.Json source generator.
// Start by defining a partial class which inherits from `WorksheetRowContext`.
// Indicate which type we want to create a row for by using the `WorksheetRow` attribute.
// The source generator will then augment this class to include the necessary code.
[WorksheetRow(typeof(Person))]
public partial class PersonRowContext : WorksheetRowContext
{
}


public static class SourceGenerator
{
    public static async Task Sample()
    {
        await using var stream = File.Create("source-generator.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        var person = new Person
        {
            FirstName = "Ola",
            LastName = "Nordmann",
            Age = 25
        };

        // Call the `AddAsRowAsync` method with the object and the matching context metadata type.
        // This will add a row to the current worksheet, with one cell per object property value.
        await spreadsheet.AddAsRowAsync(person, PersonRowContext.Default.Person);

        await spreadsheet.FinishAsync();
    }
}
