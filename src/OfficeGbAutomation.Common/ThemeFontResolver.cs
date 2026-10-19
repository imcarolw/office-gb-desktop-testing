using System.Xml.Linq;

namespace OfficeGbAutomation.Common;

public static class ThemeFontResolver
{
    public static string Resolve(string font, string themePath, string script)
    {
        if (string.IsNullOrWhiteSpace(font))
            throw new InvalidDataException("Office did not expose an East Asian font or theme token.");
        if (!font.StartsWith("+", StringComparison.Ordinal))
            return font;
        var family = font.StartsWith("+mj-", StringComparison.Ordinal) ? "majorFont"
            : font.StartsWith("+mn-", StringComparison.Ordinal) ? "minorFont"
            : throw new InvalidDataException($"Unsupported theme font token '{font}'.");
        var xml = XDocument.Load(themePath);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var group = xml.Descendants(a + family).Single();
        var specific = group.Elements(a + "font")
            .SingleOrDefault(e => (string?)e.Attribute("script") == script);
        var resolved = (string?)specific?.Attribute("typeface")
            ?? (string?)group.Element(a + "ea")?.Attribute("typeface");
        if (string.IsNullOrWhiteSpace(resolved))
            throw new InvalidDataException($"Theme token '{font}' has no font for script '{script}'.");
        return resolved;
    }
}
