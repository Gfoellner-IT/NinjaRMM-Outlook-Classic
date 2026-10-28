using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace NinjaRmm.Outlook.Infrastructure
{
    public sealed class LocalSettings
    {
        private readonly string path;
        public LocalSettings(string path = null)
        {
            this.path = path ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "NinjaRMM-Outlook-Classic", "settings.ini");
        }
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileStringW", CharSet = CharSet.Unicode)]
        private static extern uint ReadIni(string section, string key, string fallback, StringBuilder value, uint size, string path);
        public string Get(string key, string fallback)
        {
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("A settings key is required.");
            var text = new StringBuilder(32768);
            uint length = ReadIni("Settings", key, fallback ?? "", text, (uint)text.Capacity, path);
            if (length >= text.Capacity - 1) throw new InvalidOperationException("Local setting exceeds the supported length.");
            return text.ToString();
        }
        public Uri ApiBase
        {
            get
            {
                Uri uri;
                if (!Uri.TryCreate(Get("ApiBaseUrl", "https://tenant.example.invalid/v2/"), UriKind.Absolute, out uri))
                    throw new InvalidOperationException("API base is invalid.");
                ClientCredentialsTokenSource.ValidateHttps(uri);
                if (!uri.AbsolutePath.EndsWith("/", StringComparison.Ordinal)) throw new InvalidOperationException("API base must end with a slash.");
                return uri;
            }
        }
        public string CredentialTarget { get { return Get("CredentialTarget", "NinjaRMM-Outlook-Classic/API"); } }
        public int AllTicketsBoard
        {
            get
            {
                int value;
                if (!int.TryParse(Get("AllTicketsBoard", "2"), out value) || value <= 0)
                    throw new InvalidOperationException("AllTicketsBoard must be a positive integer.");
                return value;
            }
        }
    }
}
