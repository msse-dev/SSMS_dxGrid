using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelGrid.Ssms;

internal static class SsmsTheme
{
    public static bool IsDark(Control fallbackControl)
    {
        var themeBackground = TryGetSsmsThemeBackground();
        if (themeBackground.HasValue)
            return IsDark(themeBackground.Value);

        for (Control? current = fallbackControl; current != null; current = current.Parent)
        {
            var color = current.BackColor;
            if (!color.IsEmpty && color.A != 0)
                return IsDark(color);
        }

        return IsDark(SystemColors.Window);
    }

    internal static bool IsDark(Color color)
    {
        // Perceived luminance is more reliable than HSL brightness for tinted themes.
        var luminance = (0.2126 * color.R) + (0.7152 * color.G) + (0.0722 * color.B);
        return luminance < 128;
    }

    private static Color? TryGetSsmsThemeBackground()
    {
        try
        {
            var shell = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(assembly => assembly.GetName().Name == "Microsoft.VisualStudio.Shell.15.0");
            var environmentColors = shell?.GetType("Microsoft.VisualStudio.PlatformUI.EnvironmentColors", false)
                ?? Type.GetType("Microsoft.VisualStudio.PlatformUI.EnvironmentColors, Microsoft.VisualStudio.Shell.15.0", false);
            var applicationType = Type.GetType("System.Windows.Application, PresentationFramework", false);
            var application = applicationType?.GetProperty("Current", BindingFlags.Public | BindingFlags.Static)?.GetValue(null, null);
            var key = environmentColors?.GetProperty("ToolWindowBackgroundColorKey", BindingFlags.Public | BindingFlags.Static)?.GetValue(null, null)
                ?? environmentColors?.GetProperty("ToolWindowBackgroundBrushKey", BindingFlags.Public | BindingFlags.Static)?.GetValue(null, null);
            if (application == null || key == null) return null;

            var resource = applicationType!.GetMethod("TryFindResource", new[] { typeof(object) })?.Invoke(application, new[] { key });
            return ReadMediaColor(resource);
        }
        catch
        {
            // Theme services are not present in tests and may be unavailable while SSMS is shutting down.
            return null;
        }
    }

    private static Color? ReadMediaColor(object? resource)
    {
        if (resource == null) return null;
        var value = resource;
        var colorProperty = resource.GetType().GetProperty("Color", BindingFlags.Public | BindingFlags.Instance);
        if (colorProperty != null) value = colorProperty.GetValue(resource, null);
        if (value == null) return null;

        var type = value.GetType();
        var red = type.GetProperty("R")?.GetValue(value, null);
        var green = type.GetProperty("G")?.GetValue(value, null);
        var blue = type.GetProperty("B")?.GetValue(value, null);
        if (red is byte r && green is byte g && blue is byte b)
            return Color.FromArgb(r, g, b);
        return null;
    }
}
