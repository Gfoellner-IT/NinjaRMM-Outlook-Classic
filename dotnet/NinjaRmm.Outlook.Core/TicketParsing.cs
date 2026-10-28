using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace NinjaRmm.Outlook.Core
{
    // Defaults mirror TicketConfig.bas. Configurable patterns stay at the boundary.
    public static class TicketParsing
    {
        public const string SubjectPattern = @"^\[example\]\s*\(#([0-9]+)\)(.*)";
        public const string ClosedMailPattern = @"Status:\s?.*?\s?Geschlossen";
        private const string TicketPrefix = @"TICKET \(B\u00fcro - ([A-Z]+)\)( / )";

        public static bool TrySubject(string subject, out string ticketNumber, out string remainingText)
        {
            ticketNumber = remainingText = string.Empty;
            Match match = Regex.Match(subject ?? string.Empty, SubjectPattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            if (!match.Success) return false;
            ticketNumber = "#" + match.Groups[1].Value;
            remainingText = match.Groups[2].Value;
            return true;
        }

        public static string FolderName(string ticketNumber, string remainingText)
        {
            string text = remainingText ?? string.Empty;
            Match match = Regex.Match(text, TicketPrefix, RegexOptions.IgnoreCase |
                RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            if (match.Success)
                text = match.Groups[1].Value + match.Groups[2].Value + text.Replace(match.Value, "").Trim(' ');
            return ticketNumber + " (" + text.Trim(' ') + ")";
        }

        public static int TicketIdFromFolderName(string folderName)
        {
            if (string.IsNullOrEmpty(folderName) || folderName[0] != '#') return 0;
            int space = folderName.IndexOf(' ');
            string token = space < 0 ? folderName.Substring(1) : folderName.Substring(1, space - 1);
            if (token.Length == 0 || token.Length > 9) return 0;
            foreach (char c in token) if (c < '0' || c > '9') return 0;
            return int.Parse(token, CultureInfo.InvariantCulture);
        }

        public static bool FolderMatchesTicket(string folderName, string ticketNumber)
        {
            return !string.IsNullOrEmpty(ticketNumber) && folderName != null &&
                (folderName == ticketNumber || folderName.StartsWith(ticketNumber + " ", StringComparison.Ordinal));
        }

        public static bool IsClosedMail(string body)
        {
            return Regex.IsMatch(body ?? string.Empty, ClosedMailPattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        }

        // The caller supplies local time and culture. Null replaces VBA's ambiguous date value 0.
        public static string ArchiveYear(DateTime? closedLocalDate, DateTime todayLocal)
        {
            return (closedLocalDate ?? todayLocal).Year.ToString(CultureInfo.InvariantCulture);
        }

        public static string ArchiveMonth(DateTime? closedLocalDate, DateTime todayLocal, CultureInfo culture)
        {
            if (culture == null) throw new ArgumentNullException("culture");
            return (closedLocalDate ?? todayLocal).ToString("MM - MMMM", culture);
        }
    }
}
