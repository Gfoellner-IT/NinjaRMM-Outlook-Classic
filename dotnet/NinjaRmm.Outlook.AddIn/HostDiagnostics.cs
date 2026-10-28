using System;
using System.Globalization;
using System.IO;

namespace NinjaRmm.Outlook.AddIn
{
    internal static class HostDiagnostics
    {
        internal static void Record(string marker)
        {
            // Fixed diagnostic markers only: no mail content, profile identifiers or credentials.
            try
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "NinjaRMM-Outlook-Classic", "vsto");
                Directory.CreateDirectory(folder);
                File.AppendAllText(Path.Combine(folder, "host-diagnostics.log"),
                    DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) + " " + marker + Environment.NewLine);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
