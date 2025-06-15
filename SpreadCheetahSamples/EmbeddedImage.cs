using SpreadCheetah;
using SpreadCheetah.Images;

namespace SpreadCheetahSamples;

public static class EmbeddedImage
{
    public static async Task Sample()
    {
        await using var stream = File.Create("embedded-image.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await using var imageStream = File.OpenRead("Assets/icon-package.png");
        var embeddedImage = await spreadsheet.EmbedImageAsync(imageStream);

        await spreadsheet.StartWorksheetAsync("Sheet 1");

        var canvas = ImageCanvas.OriginalSize("C3");
        spreadsheet.AddImage(canvas, embeddedImage);

        await spreadsheet.FinishAsync();
    }
}
