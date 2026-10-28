using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace NinjaRmm.Outlook.Core
{
    public enum MatchOutcome { None, Found, Ambiguous }

    public sealed class TicketFacts
    {
        public int Id { get; private set; }
        public string Summary { get; private set; }
        public bool IsOpen { get; private set; }
        public TicketFacts(int id, string summary, bool isOpen)
        {
            if (id <= 0) throw new ArgumentOutOfRangeException("id");
            Id = id; Summary = summary ?? ""; IsOpen = isOpen;
        }
    }

    public sealed class MatchResult
    {
        public MatchOutcome Outcome { get; private set; }
        public MatchRule Rule { get; private set; }
        public int TicketId { get; private set; }
        internal MatchResult(MatchOutcome outcome, MatchRule rule, int id = 0)
        { Outcome = outcome; Rule = rule; TicketId = id; }
    }

    // Infrastructure supplies a snapshot; matching itself never accesses Outlook or the API.
    public interface ITicketIndex
    {
        TicketFacts GetTicket(int id);
        IEnumerable<TicketFacts> FindByExternalNumber(string number);
        IEnumerable<TicketFacts> FindOpenBySubject(string normalizedSubject);
        IEnumerable<TicketFacts> FindOpenByStem(string stem);
        int FindByConversation(string conversationId);
    }

    public sealed class LearnedRule
    {
        public string Domain { get; private set; }
        public string Stem { get; private set; }
        public bool Enabled { get; private set; }
        public LearnedRule(string domain, string stem, bool enabled = false)
        {
            Domain = (domain ?? "").Trim().ToLowerInvariant();
            Stem = (stem ?? "").Trim().ToLowerInvariant(); Enabled = enabled;
        }
    }

    public sealed class TicketMatcher
    {
        private readonly MatchLevels levels;
        private readonly HashSet<string> excluded;
        private readonly LearnedRule[] learned;
        private readonly string ticketPrefix, organizationPrefix;
        private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
        private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(1);
        private const string NinjaPattern = @"(?:ticket\s*-?\s*(?:nr\.?\s*)?#?\s*|/ticket/|\(#)([0-9]{4,6})(?![0-9])";
        private const string ExternalPattern = @"\[(?:CSM-?)?TicketNr\.?\s*([0-9]{6,12})\]";

        public TicketMatcher(MatchLevels levels = null, string excludedDomains = "github.com,newsletter.example",
            IEnumerable<LearnedRule> learnedRules = null,
            string ticketPrefix = @"TICKET \(B\u00fcro - ([A-Z]+)\)( / )",
            string organizationPrefix = @"^\s*B.{1,2}ro\s*:?\s*-?\s*")
        {
            this.levels = levels ?? new MatchLevels();
            excluded = new HashSet<string>((excludedDomains ?? "").Split(',')
                .Select(x => x.Trim()).Where(x => x.Length > 0), StringComparer.OrdinalIgnoreCase);
            learned = (learnedRules ?? Enumerable.Empty<LearnedRule>()).Where(x => x != null).ToArray();
            this.ticketPrefix = ticketPrefix; this.organizationPrefix = organizationPrefix;
        }

        private static string Replace(string text, string pattern, string replacement)
        { return Regex.Replace(text, pattern, replacement, Options, Timeout); }

        public string NormalizeSubject(string text)
        {
            string s = Replace(text ?? "", @"\[(?:CSM-?)?TicketNr\.?\s*[0-9]+\]", " ");
            s = Replace(s, @"^\s*(?:(?:AW|RE|WG|FW|FWD|Antwort|Weiterleitung)\s*:\s*)+", "");
            s = Replace(s, @"^\s*(?:CSM|CMS)\s*\|\s*", "");
            s = Replace(s, @"^[\s\-:|]+", "");
            s = Replace(s, @"^\s*" + ticketPrefix, "");
            s = Replace(s, @"^\s*(?:(?:AW|RE|WG|FW|FWD)\s*:\s*)+", "");
            s = Replace(s, organizationPrefix, "");
            s = Replace(s, @"^[\s\-:|]+", "");
            return Replace(s, @"\s+", " ").Trim().ToLowerInvariant();
        }

        public static string SubjectStem(string normalized, int maxWords = 4)
        {
            if (maxWords < 1) throw new ArgumentOutOfRangeException("maxWords");
            string s = Replace(Replace(normalized ?? "", @"[0-9]+", "#"), " +", " ").Trim();
            return string.Join(" ", s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Take(maxWords));
        }

        public static bool StemStartsWith(string stem, string prefix)
        {
            string s = (stem ?? "").Trim(), p = (prefix ?? "").Trim();
            return p.Length > 0 && (s.Equals(p, StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith(p + " ", StringComparison.OrdinalIgnoreCase));
        }

        public static IList<int> ExtractNinjaIds(string text)
        { return Extract(text, NinjaPattern).Select(x => int.Parse(x, CultureInfo.InvariantCulture)).Distinct().ToArray(); }
        public static IList<string> ExtractExternalNumbers(string text) { return Extract(text, ExternalPattern); }
        private static IList<string> Extract(string text, string pattern)
        {
            return Regex.Matches(text ?? "", pattern, Options, Timeout).Cast<Match>()
                .Select(x => x.Groups[1].Value).Distinct(StringComparer.Ordinal).ToArray();
        }
        public static bool HasAutoHeader(string headers)
        {
            return Regex.IsMatch(headers ?? "", @"^(?:X-Auto-Response-Suppress|List-Id|List-Unsubscribe|Precedence\s*:\s*(?:bulk|list|junk)|Auto-Submitted\s*:\s*auto-)",
                Options | RegexOptions.Multiline, Timeout);
        }
        public bool IsExcluded(string domain, bool autoHeader)
        { return autoHeader || excluded.Contains((domain ?? "").Trim()); }
        public static bool IsSpecificSubject(string normalized)
        { return !string.IsNullOrEmpty(normalized) && normalized.Length >= 20 && normalized.Split(' ').Length >= 3; }

        public MatchResult Match(string subject, string body, string conversation, string domain, bool autoHeader, ITicketIndex index)
        {
            if (index == null || !levels.AnyActive || IsExcluded(domain, autoHeader)) return None();
            var ids = levels.For(MatchRule.NinjaId) == MatchLevel.Off ? new int[0] : ExtractNinjaIds(subject);
            if (ids.Count == 0 && levels.For(MatchRule.NinjaId) != MatchLevel.Off) ids = ExtractNinjaIds(body);
            if (ids.Count > 1) return Ambiguous(MatchRule.NinjaId);
            if (ids.Count == 1)
            {
                var ticket = index.GetTicket(ids[0]);
                if (ticket != null) return Found(MatchRule.NinjaId, ticket.Id);
            }
            var numbers = levels.For(MatchRule.ExternalNumber) == MatchLevel.Off ? new string[0] : ExtractExternalNumbers(subject);
            if (numbers.Count > 1) return Ambiguous(MatchRule.ExternalNumber);
            if (numbers.Count == 1)
            {
                var all = (index.FindByExternalNumber(numbers[0]) ?? Enumerable.Empty<TicketFacts>()).Where(x => x != null).ToArray();
                var open = all.Where(x => x.IsOpen).ToArray();
                if (all.Length == 1) return Found(MatchRule.ExternalNumber, all[0].Id);
                if (open.Length == 1) return Found(MatchRule.ExternalNumber, open[0].Id);
                return all.Length > 1 ? Ambiguous(MatchRule.ExternalNumber) : None(MatchRule.ExternalNumber);
            }
            if (!string.IsNullOrEmpty(conversation) && levels.For(MatchRule.Conversation) != MatchLevel.Off)
            {
                int id = index.FindByConversation(conversation);
                if (id > 0) return Found(MatchRule.Conversation, id);
            }
            string normalized = NormalizeSubject(subject);
            if (levels.For(MatchRule.Subject) != MatchLevel.Off && IsSpecificSubject(normalized))
            {
                var matches = (index.FindOpenBySubject(normalized) ?? Enumerable.Empty<TicketFacts>())
                    .Where(t => t != null && t.IsOpen && NormalizeSubject(t.Summary) == normalized).ToArray();
                if (matches.Length == 1) return Found(MatchRule.Subject, matches[0].Id);
                if (matches.Length > 1) return Ambiguous(MatchRule.Subject);
            }
            if (levels.For(MatchRule.Learned) != MatchLevel.Off)
            {
                string stem = SubjectStem(normalized);
                var rule = learned.FirstOrDefault(r => r.Enabled && r.Stem.Length >= 8 && r.Domain.Length > 0 &&
                    (r.Domain == "*" || r.Domain.Equals((domain ?? "").Trim(), StringComparison.OrdinalIgnoreCase)) && StemStartsWith(stem, r.Stem));
                if (rule != null)
                {
                    var matches = (index.FindOpenByStem(rule.Stem) ?? Enumerable.Empty<TicketFacts>())
                        .Where(t => t != null && t.IsOpen && StemStartsWith(SubjectStem(NormalizeSubject(t.Summary)), rule.Stem)).ToArray();
                    if (matches.Length == 1) return Found(MatchRule.Learned, matches[0].Id);
                    if (matches.Length > 1) return Ambiguous(MatchRule.Learned);
                }
            }
            return None();
        }
        private static MatchResult None(MatchRule rule = MatchRule.None) { return new MatchResult(MatchOutcome.None, rule); }
        private static MatchResult Found(MatchRule rule, int id) { return new MatchResult(MatchOutcome.Found, rule, id); }
        private static MatchResult Ambiguous(MatchRule rule) { return new MatchResult(MatchOutcome.Ambiguous, rule); }
    }
}
