using System;
using System.Collections.Generic;
using System.Linq;
using NinjaRmm.Outlook.Core;

namespace NinjaRmm.Outlook.Core.Tests
{
    internal static partial class Program
    {
        private static void RunMatcherTests()
        {
            Run("Matcher_ExtractionAndNormalization", MatcherParsing);
            Run("Matcher_PriorityAndExternalStop", MatcherPriority);
            Run("Matcher_ExclusionsAndDisabledRulesAvoidLookups", MatcherExclusions);
            Run("Matcher_SubjectAndLearnedAmbiguity", MatcherSubjects);
            Run("Matcher_ClosedUnreadableAndStaleConfirmation", MatcherSafety);
            Run("Matcher_DryRunAndLevelMatrix", MatcherActions);
            Run("Time_UtcMonthBoundaryAndDst", TimeZones);
        }

        private sealed class MemoryIndex : ITicketIndex
        {
            internal readonly List<TicketFacts> Tickets = new List<TicketFacts>();
            internal readonly Dictionary<string, int> Conversations = new Dictionary<string, int>();
            internal int Calls;
            public TicketFacts GetTicket(int id) { Calls++; return Tickets.FirstOrDefault(t => t.Id == id); }
            public IEnumerable<TicketFacts> FindByExternalNumber(string n) { Calls++; return Tickets.Where(t => t.Summary.Contains(n)); }
            public IEnumerable<TicketFacts> FindOpenBySubject(string s) { Calls++; return Tickets; } // Deliberately loose; matcher must recheck.
            public IEnumerable<TicketFacts> FindOpenByStem(string s) { Calls++; return Tickets; }
            public int FindByConversation(string c) { Calls++; int id; return Conversations.TryGetValue(c, out id) ? id : 0; }
        }
        private static MemoryIndex ExampleIndex()
        {
            var index = new MemoryIndex();
            index.Tickets.Add(new TicketFacts(20103, "CSM | [TicketNr. 900000105] - Example invoice processing problem", true));
            index.Tickets.Add(new TicketFacts(20108, "Example warehouse configuration request", true));
            index.Tickets.Add(new TicketFacts(7001, "[TicketNr. 900000107] Previous closed request", false));
            index.Conversations.Add("conversation-a", 20108);
            return index;
        }
        private static MatchResult Match(TicketMatcher matcher, MemoryIndex index, string subject,
            string body = "", string conversation = "", string domain = "vendor.example", bool auto = false)
        { return matcher.Match(subject, body, conversation, domain, auto, index); }

        private static void MatcherParsing()
        {
            foreach (string text in new[] { "Ticket 20108", "Ticket #20108", "Ticket-Nr. 20108", "(#20108)", "https://x.example/ticket/20108" })
                Equal(20108, TicketMatcher.ExtractNinjaIds(text).Single());
            Equal(1, TicketMatcher.ExtractNinjaIds("Ticket 20108 Ticket 20108").Count);
            foreach (string text in new[] { "Ticket 12", "Ticket 1234567", "[TicketNr. 900000105]", "invoice #20108", "Ticket \u0661\u0662\u0663\u0664" })
                Equal(0, TicketMatcher.ExtractNinjaIds(text).Count);
            Equal("900000105", TicketMatcher.ExtractExternalNumbers("[CSM-TicketNr. 900000105]").Single());
            var matcher = new TicketMatcher();
            Equal("example invoice processing problem", matcher.NormalizeSubject("AW: CSM | [TicketNr. 900000105] - Example invoice processing problem"));
            Equal("example interface change request", matcher.NormalizeSubject("TICKET (B\u00fcro - AT) / Example interface change request"));
            Equal("item235 - error", matcher.NormalizeSubject("AW: [TicketNr. 900000106] B\u00fcro: item235 - error"));
            Equal("a b", matcher.NormalizeSubject(" RE:   a  b "));
            Equal("item# - error #", TicketMatcher.SubjectStem("item235 - error 4711 in form220"));
            Equal(false, TicketMatcher.StemStartsWith("item# - error", "item"));
            Equal(true, TicketMatcher.StemStartsWith("item# - error #", "ITEM# - error"));
            Throws<ArgumentOutOfRangeException>(() => TicketMatcher.SubjectStem("text", 0));
            foreach (string header in new[] { "List-Id: a", "X-Auto-Response-Suppress: All", "X: a\r\nPrecedence: bulk", "Auto-Submitted: auto-replied" })
                Equal(true, TicketMatcher.HasAutoHeader(header));
            Equal(false, TicketMatcher.HasAutoHeader("Subject: Auto-Submitted: auto-replied"));
        }
        private static void MatcherPriority()
        {
            var matcher = new TicketMatcher(); var index = ExampleIndex();
            Equal(20108, Match(matcher, index, "Ticket 20108 [TicketNr. 900000105]").TicketId);
            Equal(20103, Match(matcher, index, "[TicketNr. 900000105]", conversation: "conversation-a").TicketId);
            var missing = Match(matcher, index, "[TicketNr. 900000199]", "Ticket 1234", "conversation-a");
            Equal(MatchOutcome.None, missing.Outcome); Equal(MatchRule.ExternalNumber, missing.Rule);
            Equal(MatchRule.Conversation, Match(matcher, index, "Example invoice processing problem", conversation: "conversation-a").Rule);
            Equal(MatchOutcome.Ambiguous, Match(matcher, index, "Ticket 20108 Ticket 20103").Outcome);
            Equal(MatchOutcome.Ambiguous, Match(matcher, index, "[TicketNr. 900000105] [TicketNr. 900000107]").Outcome);
            Equal(20108, Match(matcher, index, "Unrelated", body: "Ticket 20108").TicketId);
            Equal(20103, Match(matcher, index, "Ticket 20103", body: "Ticket 20108 Ticket 7001").TicketId);
            index.Tickets.Add(new TicketFacts(7002, "[TicketNr. 900000107] Current reopened request", true));
            Equal(7002, Match(matcher, index, "[TicketNr. 900000107]").TicketId);
            index.Tickets.Add(new TicketFacts(7003, "[TicketNr. 900000107] Another open request", true));
            Equal(MatchOutcome.Ambiguous, Match(matcher, index, "[TicketNr. 900000107]").Outcome);
        }
        private static void MatcherExclusions()
        {
            var index = ExampleIndex(); var matcher = new TicketMatcher();
            Equal(MatchOutcome.None, Match(matcher, index, "Ticket 20108", domain: " GitHub.COM ").Outcome);
            Equal(MatchOutcome.None, Match(matcher, index, "Ticket 20108", auto: true).Outcome);
            Equal(0, index.Calls);
            Equal(MatchOutcome.None, Match(new TicketMatcher(new MatchLevels(cap: 0)), index, "Ticket 20108").Outcome);
            Equal(0, index.Calls);
            Equal(20103, Match(new TicketMatcher(new MatchLevels(ninjaId: 0)), index, "Ticket 20108 [TicketNr. 900000105]").TicketId);
            Equal(MatchOutcome.None, matcher.Match("Ticket 20108", "", "", "", false, null).Outcome);
        }
        private static void MatcherSubjects()
        {
            var index = ExampleIndex(); var matcher = new TicketMatcher();
            Equal(MatchRule.Subject, Match(matcher, index, "AW: Example warehouse configuration request").Rule);
            index.Tickets.Add(new TicketFacts(20109, "Example warehouse configuration request", false));
            Equal(20108, Match(matcher, index, "Example warehouse configuration request").TicketId);
            index.Tickets.Add(new TicketFacts(20110, "Example warehouse configuration request", true));
            Equal(MatchOutcome.Ambiguous, Match(matcher, index, "Example warehouse configuration request").Outcome);
            Equal(false, TicketMatcher.IsSpecificSubject("onewordlongenoughtobeaticket"));
            index = new MemoryIndex(); index.Tickets.Add(new TicketFacts(20111, "item235 - error 4711 in form220", true));
            var rule = new LearnedRule("vendor.example", "item# - error", true);
            matcher = new TicketMatcher(learnedRules: new[] { rule });
            Equal(MatchRule.Learned, Match(matcher, index, "item999 - error 123 in form4").Rule);
            Equal(MatchOutcome.None, Match(matcher, index, "item999 - error 123", domain: "other.example").Outcome);
            Equal(MatchOutcome.None, Match(new TicketMatcher(learnedRules: new[] { new LearnedRule("vendor.example", "item# - error") }), index, "item999 - error 123").Outcome);
            index.Tickets.Add(new TicketFacts(20112, "item777 - error 888 in form1", true));
            Equal(MatchOutcome.Ambiguous, Match(matcher, index, "item999 - error 123").Outcome);
        }
        private static void MatcherSafety()
        {
            var index = ExampleIndex(); var result = Match(new TicketMatcher(), index, "Ticket 7001");
            Equal(MatchOutcome.Found, result.Outcome); // Recognition is distinct from permission to move.
            Equal(MatchAction.None, MatchingPolicy.Decide(result, new MatchLevels(ninjaId: 3), false, index.GetTicket(7001), false, true));
            Equal(false, MatchingPolicy.CanConfirm(7001, index.GetTicket(7001), true));
            Equal(false, MatchingPolicy.CanConfirm(20108, null, true));
            Equal(false, MatchingPolicy.CanConfirm(20108, index.GetTicket(20103), true));
            Equal(false, MatchingPolicy.CanConfirm(20108, index.GetTicket(20108), false));
            Equal(true, MatchingPolicy.CanConfirm(20108, index.GetTicket(20108), true));
            Equal(false, MatchingPolicy.CanConfirm(20108, new TicketFacts(20108, "Now closed", false), true));
        }
        private static void MatcherActions()
        {
            var index = ExampleIndex(); var result = Match(new TicketMatcher(), index, "Ticket 20108");
            MatchAction[] expected = { MatchAction.None, MatchAction.LogOnly, MatchAction.Propose, MatchAction.Move };
            for (int level = 0; level <= 3; level++)
            {
                var levels = new MatchLevels(ninjaId: level);
                Equal(expected[level], MatchingPolicy.Decide(result, levels, false, index.GetTicket(20108), false, true));
                Equal(level == 0 ? MatchAction.None : MatchAction.LogOnly,
                    MatchingPolicy.Decide(result, levels, true, index.GetTicket(20108), false, true));
            }
            Equal(MatchAction.None, MatchingPolicy.Decide(result, new MatchLevels(ninjaId: 2), false, index.GetTicket(20108), true, true));
            Equal(MatchAction.None, MatchingPolicy.Decide(result, new MatchLevels(ninjaId: 3), false, null, false, true));
            Equal(MatchAction.None, MatchingPolicy.Decide(result, new MatchLevels(ninjaId: 3), false, index.GetTicket(20103), false, true));
            Equal(MatchAction.None, MatchingPolicy.Decide(result, new MatchLevels(ninjaId: 3), false, index.GetTicket(20108), false, false));
        }
        private static void TimeZones()
        {
            var zone = TimeZoneInfo.FindSystemTimeZoneById("W. Europe Standard Time");
            Equal(new DateTime(2026, 11, 1, 0, 30, 0), TicketTime.UtcToLocal(new DateTime(2026, 10, 31, 23, 30, 0, DateTimeKind.Utc), zone).Value);
            Equal(new DateTime(2026, 7, 1, 2, 0, 0), TicketTime.UtcToLocal(new DateTime(2026, 7, 1), zone).Value);
            Equal(new DateTime(2026, 1, 1, 1, 0, 0), TicketTime.UtcToLocal(new DateTime(2026, 1, 1), zone).Value);
            Equal(false, TicketTime.UtcToLocal(null, zone).HasValue);
            Throws<ArgumentException>(() => TicketTime.UtcToLocal(DateTime.SpecifyKind(new DateTime(2026, 1, 1), DateTimeKind.Local), zone));
        }
    }
}
