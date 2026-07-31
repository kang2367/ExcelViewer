using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Vestris.ResourceLib;

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: IconPatcher <exePath> <icoPath>");
    return 1;
}

var exePath = Path.GetFullPath(args[0]);
var icoPath = Path.GetFullPath(args[1]);

if (!File.Exists(exePath))
{
    Console.Error.WriteLine($"Executable not found: {exePath}");
    return 1;
}

if (!File.Exists(icoPath))
{
    Console.Error.WriteLine($"Icon not found: {icoPath}");
    return 1;
}

var tempIcoPath = Path.Combine(Path.GetTempPath(), $"iconpatch_{Guid.NewGuid():N}.ico");
try
{
    File.WriteAllBytes(tempIcoPath, MultiSizeIconBuilder.BuildFromIco(icoPath));
    var iconFile = new IconFile(tempIcoPath);

    SaveIconGroup(exePath, iconFile, 1);
    SaveIconGroup(exePath, iconFile, 32512);
}
finally
{
    if (File.Exists(tempIcoPath))
    {
        File.Delete(tempIcoPath);
    }
}

Console.WriteLine($"Updated icon: {exePath}");
return 0;

static void SaveIconGroup(string exePath, IconFile iconFile, int resourceId)
{
    var iconDirectoryResource = new IconDirectoryResource(iconFile)
    {
        Name = new ResourceId(resourceId),
        Language = ResourceUtil.NEUTRALLANGID
    };
    iconDirectoryResource.SaveTo(exePath);
}

internal static class MultiSizeIconBuilder
{
    private static readonly int[] Sizes = [16, 32, 48, 256];

    public static byte[] BuildFromIco(string sourceIcoPath)
    {
        using var source = new Icon(sourceIcoPath);
        using var sourceBitmap = source.ToBitmap();
        var images = Sizes.Select(size => (Size: size, PngData: CreateImageData(sourceBitmap, size))).ToArray();
        return BuildIco(images);
    }

    private static byte[] CreateImageData(Bitmap source, int size)
    {
        using var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.Transparent);
            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            graphics.DrawImage(source, new Rectangle(0, 0, size, size));
        }

        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private static byte[] BuildIco((int Size, byte[] PngData)[] images)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream);

        writer.Write((ushort)0);
        writer.Write((ushort)1);
        writer.Write((ushort)images.Length);

        var offset = 6 + images.Length * 16;
        foreach (var (size, pngData) in images)
        {
            writer.Write((byte)(size >= 256 ? 0 : size));
            writer.Write((byte)(size >= 256 ? 0 : size));
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((ushort)1);
            writer.Write((ushort)32);
            writer.Write((uint)pngData.Length);
            writer.Write((uint)offset);
            offset += pngData.Length;
        }

        foreach (var (_, pngData) in images)
        {
            writer.Write(pngData);
        }

        return stream.ToArray();
    }
}
