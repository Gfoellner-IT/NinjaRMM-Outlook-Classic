using System;
using System.Globalization;
using NinjaRmm.Outlook.Core;

namespace NinjaRmm.Outlook.Core.Tests
{
    // No package restore, Office, credentials or live API is required.
    internal static partial class Program
    {
        private static int tests, failures, assertions;

        private static int Main()
        {
            Run("Archive_AllStateMailAndFallbackCombinations", ArchiveMatrix);
            Run("Archive_InvalidStateIsRejected", ArchiveInvalidState);
            Run("Levels_DefaultsAndOff", LevelDefaults);
            Run("Levels_AllRulesRespectCapAndClamping", LevelMatrix);
            Run("Levels_IndependentInstances", IndependentLevels);
            Run("Parsing_SubjectReferenceCases", Subjects);
            Run("Parsing_FolderReferenceCases", Folders);
            Run("Parsing_ClosedMailReferenceCases", ClosedMail);
            Run("Archive_LocalDateCultureAndFallback", ArchiveDates);
            RunMatcherTests();
            RunResponseTests();
            Console.WriteLine("Tests passed: {0}, failed: {1}, assertions: {2}", tests - failures, failures, assertions);
            Console.WriteLine(failures == 0 ? "RESULT: OK" : "RESULT: FAILED");
            return failures == 0 ? 0 : 1;
        }

        private static void Run(string name, Action test)
        {
            tests++;
            try { test(); Console.WriteLine("PASS " + name); }
            catch (Exception ex) { failures++; Console.WriteLine("FAIL " + name + ": " + ex.Message); }
        }

        private static void Equal<T>(T expected, T actual)
        {
            assertions++;
            if (!object.Equals(expected, actual))
                throw new InvalidOperationException("Expected " + expected + ", got " + actual);
        }

        private static void Throws<T>(Action action) where T : Exception
        {
            assertions++;
            try { action(); } catch (T) { return; }
            throw new InvalidOperationException("Expected " + typeof(T).Name);
        }

        private static void ArchiveMatrix()
        {
            // Explicit behavior table from TicketDecisionTests.bas, not the production algorithm.
            TicketState[] states = { TicketState.Unavailable, TicketState.Open, TicketState.Closed, TicketState.NotFound };
            ArchiveDecision[,] expected = {
                { ArchiveDecision.Keep, ArchiveDecision.Keep, ArchiveDecision.Keep, ArchiveDecision.ByMailFallback },
                { ArchiveDecision.Keep, ArchiveDecision.Keep, ArchiveDecision.Keep, ArchiveDecision.Keep },
                { ArchiveDecision.ByApi, ArchiveDecision.ByApi, ArchiveDecision.ByApi, ArchiveDecision.ByApi },
                { ArchiveDecision.ByApi, ArchiveDecision.ByApi, ArchiveDecision.ByApi, ArchiveDecision.ByApi }
            };
            for (int row = 0; row < states.Length; row++)
            {
                Equal(row == 0, ArchivePolicy.NeedsMailCheck(states[row], true));
                Equal(false, ArchivePolicy.NeedsMailCheck(states[row], false));
                for (int flags = 0; flags < 4; flags++)
                    Equal(expected[row, flags], ArchivePolicy.Decide(states[row], (flags & 2) != 0, (flags & 1) != 0));
            }
        }

        private static void ArchiveInvalidState()
        {
            Throws<ArgumentOutOfRangeException>(() => ArchivePolicy.Decide((TicketState)99, true, true));
            Throws<ArgumentOutOfRangeException>(() => ArchivePolicy.NeedsMailCheck((TicketState)(-1), true));
        }

        private static void LevelDefaults()
        {
            var defaults = new MatchLevels();
            foreach (MatchRule rule in new[] { MatchRule.NinjaId, MatchRule.ExternalNumber,
                MatchRule.Conversation, MatchRule.Subject, MatchRule.Learned })
                Equal(MatchLevel.LogOnly, defaults.For(rule));
            Equal(MatchLevel.Off, defaults.For(MatchRule.None));
            Equal(MatchLevel.Off, defaults.For((MatchRule)99));
            Equal(true, defaults.AnyActive);
            Equal(true, defaults.AnyAt(MatchLevel.LogOnly));
            Equal(false, defaults.AnyAt(MatchLevel.Confirm));
            Equal(false, new MatchLevels(cap: 0).AnyActive);
            Equal(false, defaults.AnyAt((MatchLevel)99));
        }

        private static void LevelMatrix()
        {
            // Rows are configured levels -1..4, columns are global caps -1..4.
            int[,] expected = {
                {0,0,0,0,0,0}, {0,0,0,0,0,0}, {0,0,1,1,1,1},
                {0,0,1,2,2,2}, {0,0,1,2,3,3}, {0,0,1,2,3,3}
            };
            for (int configured = -1; configured <= 4; configured++)
                for (int cap = -1; cap <= 4; cap++)
                {
                    var levels = new MatchLevels(configured, configured, configured, configured, configured, cap);
                    foreach (MatchRule rule in new[] { MatchRule.NinjaId, MatchRule.ExternalNumber,
                        MatchRule.Conversation, MatchRule.Subject, MatchRule.Learned })
                        Equal((MatchLevel)expected[configured + 1, cap + 1], levels.For(rule));
                }
        }

        private static void IndependentLevels()
        {
            var first = new MatchLevels(0, 1, 2, 3, 2, 3);
            Equal(MatchLevel.Off, first.For(MatchRule.NinjaId));
            Equal(MatchLevel.LogOnly, first.For(MatchRule.ExternalNumber));
            Equal(MatchLevel.Confirm, first.For(MatchRule.Conversation));
            Equal(MatchLevel.Automatic, first.For(MatchRule.Subject));
            Equal(MatchLevel.Confirm, first.For(MatchRule.Learned));
            Equal(MatchLevel.LogOnly, new MatchLevels().For(MatchRule.Subject));
            Equal(MatchLevel.Automatic, first.For(MatchRule.Subject));
        }

        private static void Subjects()
        {
            string number, text;
            Equal(true, TicketParsing.TrySubject("[example] (#3621) Printer broken", out number, out text));
            Equal("#3621", number); Equal(" Printer broken", text);
            Equal(true, TicketParsing.TrySubject("[example](#7)x", out number, out text));
            Equal("#7", number); Equal("x", text);
            Equal(true, TicketParsing.TrySubject("[EXAMPLE] (#12) x", out number, out text));
            Equal("#12", number);
            Equal(true, TicketParsing.TrySubject("[example] (#0012) x", out number, out text));
            Equal("#0012", number);
            foreach (string invalid in new[] { "Re: [example] (#1) x", "[other] (#1) x",
                "[example] (no number)", "[example] (#\u0661) x", "", null })
            {
                Equal(false, TicketParsing.TrySubject(invalid, out number, out text));
                Equal("", number); Equal("", text);
            }
        }

        private static void Folders()
        {
            Equal("#3621 (Printer broken)", TicketParsing.FolderName("#3621", " Printer broken"));
            Equal("#3621 ()", TicketParsing.FolderName("#3621", ""));
            Equal("#5 (IT / Printer defekt)", TicketParsing.FolderName("#5", " TICKET (B\u00fcro - IT) / Printer defekt"));
            Equal("#5 (HR / Urlaub)", TicketParsing.FolderName("#5", " TICKET (B\u00fcro - HR) / Urlaub"));
            Equal(3621, TicketParsing.TicketIdFromFolderName("#3621 (x)"));
            Equal(3621, TicketParsing.TicketIdFromFolderName("#3621"));
            Equal(999999999, TicketParsing.TicketIdFromFolderName("#999999999 (x)"));
            foreach (string invalid in new[] { "3621 (x)", "#abc (x)", "#36x (x)", "#", "#1234567890", "#\u0661", null })
                Equal(0, TicketParsing.TicketIdFromFolderName(invalid));
            Equal(true, TicketParsing.FolderMatchesTicket("#36 (x)", "#36"));
            Equal(true, TicketParsing.FolderMatchesTicket("#36", "#36"));
            Equal(false, TicketParsing.FolderMatchesTicket("#3621 (x)", "#36"));
            Equal(false, TicketParsing.FolderMatchesTicket("#36 (x)", ""));
            Equal(false, TicketParsing.FolderMatchesTicket(null, "#36"));
        }

        private static void ClosedMail()
        {
            Equal(true, TicketParsing.IsClosedMail("Status: Ticket Geschlossen"));
            Equal(true, TicketParsing.IsClosedMail("status: geschlossen"));
            Equal(false, TicketParsing.IsClosedMail("Status: Offen"));
            Equal(false, TicketParsing.IsClosedMail("Status: Closed"));
            Equal(false, TicketParsing.IsClosedMail(null));
        }

        private static void ArchiveDates()
        {
            var today = new DateTime(2026, 10, 27);
            var closed = new DateTime(2025, 4, 30);
            Equal("2025", TicketParsing.ArchiveYear(closed, today));
            Equal("2026", TicketParsing.ArchiveYear(null, today));
            Equal("04 - April", TicketParsing.ArchiveMonth(closed, today, CultureInfo.GetCultureInfo("de-AT")));
            Equal("10 - Oktober", TicketParsing.ArchiveMonth(null, today, CultureInfo.GetCultureInfo("de-AT")));
            Equal("10 - October", TicketParsing.ArchiveMonth(null, today, CultureInfo.GetCultureInfo("en-US")));
            Throws<ArgumentNullException>(() => TicketParsing.ArchiveMonth(closed, today, null));
        }
    }
}
