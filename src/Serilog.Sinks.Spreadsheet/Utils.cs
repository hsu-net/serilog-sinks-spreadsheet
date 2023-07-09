using static System.Diagnostics.Process;
// ReSharper disable MemberCanBePrivate.Global

namespace Serilog.Sinks.Spreadsheet;

internal static class Utils
{
    public static bool Exists(string file)
    {
        return File.Exists(Path.IsPathRooted(file) ? file : Path.Combine(GetDirectory(),file));
    }
    
    public static string GetDirectory()
    {
        return Path.GetDirectoryName(GetCurrentProcess().MainModule!.FileName)!;
    }
}
