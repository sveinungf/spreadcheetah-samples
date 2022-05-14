using SpreadCheetah;
using SpreadCheetah.Validations;

namespace SpreadCheetahSamples;

public static class DataValidations
{
    public static async Task Sample()
    {
        await using var stream = File.Create("data-validations.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        // Data Validations is a feature of Excel, which validate data entered into cells.
        // Note that these validations are not enforced for cell values created by SpreadCheetah.

        // Validation of integers, decimals, and text lengths supports common operators: <, >, <=, >=, =, !=
        var positiveInteger = DataValidation.IntegerGreaterThan(0);
        positiveInteger.InputTitle = "Positive integer";
        positiveInteger.InputMessage = "Enter a positive integer";
        spreadsheet.AddDataValidation("A1", positiveInteger);

        // Can reuse the same data validation for a different cell or a range of cells.
        spreadsheet.AddDataValidation("C1:F1", positiveInteger);

        // Validating ranges takes two operands.
        var decimalRange = DataValidation.DecimalBetween(100, 200);
        decimalRange.InputTitle = "Range";
        decimalRange.InputMessage = "Enter a number between 100 and 200";
        spreadsheet.AddDataValidation("A2", decimalRange);

        // Can also optionally set the error message to be shown among other properties.
        var textLengthLimit = DataValidation.TextLengthLessThanOrEqualTo(20);
        textLengthLimit.InputTitle = "Name";
        textLengthLimit.InputMessage = "Enter your name";
        textLengthLimit.ErrorTitle = "Text length limit";
        textLengthLimit.ErrorMessage = "Max 20 characters";
        textLengthLimit.ErrorType = ValidationErrorType.Warning;
        textLengthLimit.IgnoreBlank = false;
        spreadsheet.AddDataValidation("A3:D3", textLengthLimit);

        // A list of allowed values, shown a dropdown menu.
        var colors = new[] { "Red", "Green", "Blue" };
        var allowedValues = DataValidation.ListValues(colors, showDropdown: true);
        allowedValues.InputTitle = "Color";
        allowedValues.InputMessage = "Choose a color";
        spreadsheet.AddDataValidation("A4:A6", allowedValues);

        await spreadsheet.FinishAsync();
    }
}
