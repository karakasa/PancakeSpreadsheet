using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Utility;
internal static class Config
{
    public static bool HideAllComponents => _hideAllComponents ??= ShouldHideAllComponents();
    private static bool? _hideAllComponents;
    private static bool ShouldHideAllComponents()
    {
        const string FileName = "HideComponents";

        return DetermineFlagFileExistence(typeof(Config), FileName) ||
            DetermineFlagFileExistence(typeof(Grasshopper.CentralSettings), FileName);
    }

    private static bool DetermineFlagFileExistence(Type type, string filename)
    {
        var loc = type.Assembly.Location;
        if (string.IsNullOrEmpty(loc)) return false;
        return DetermineFlagFileExistence(loc, filename);
    }
    private static bool DetermineFlagFileExistence(string rootFile, string filename)
    {
        try
        {
            var path = Path.Combine(Path.GetDirectoryName(rootFile), filename);
            return File.Exists(path);
        }
        catch
        {
            return false;
        }
    }
}
