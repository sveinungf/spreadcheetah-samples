using SpreadCheetah;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.Styling;
using SpreadCheetahSamples.Helpers;

namespace SpreadCheetahSamples;

// A C# class which we want to add as a row in a worksheet.
// The source generator will pick the properties with public getters.
// By default, the order of the properties will decide the order of the cell values.
// The order can also be customized by using the ColumnOrder attribute.
public class Person
{
    public string? Title { get; set; }

    [ColumnHeader("First name")]
    public string? FirstName { get; set; }

    [ColumnHeader(typeof(HeaderResources), nameof(HeaderResources.Header_LastName))]
    public string? LastName { get; set; }

    [ColumnOrder(3)]
    public string? MiddleName { get; set; }

    public int Age { get; set; }
}


// The source generator is used in a similar way as the System.Text.Json source generator.
// Start by defining a partial class which inherits from `WorksheetRowContext`.
// Indicate which type we want to create a row for by using the `WorksheetRow` attribute.
// The source generator will then augment this class to include the necessary code.
[WorksheetRow(typeof(Person))]
public partial class PersonRowContext : WorksheetRowContext;


public static class SourceGenerator
{
    public static async Task Sample()
    {
        await using var stream = File.Create("source-generator.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        var headerStyle = new Style { Font = { Bold = true } };
        var headerStyleId = spreadsheet.AddStyle(headerStyle);

        // The 'AddHeaderRowAsync' method will add a row of header names to the worksheet.
        // By default, the property names will be used. This can be customized by using the ColumnHeader attribute.
        await spreadsheet.AddHeaderRowAsync(PersonRowContext.Default.Person, headerStyleId);

        var person = new Person
        {
            Title = "Mr.",
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Age = 25
        };

        // Call the 'AddAsRowAsync' method with the object and the context type created by the source generator.
        // This will add a row to the current worksheet, with one cell per object property value.
        await spreadsheet.AddAsRowAsync(person, PersonRowContext.Default.Person);

        await spreadsheet.FinishAsync();
    }
}
