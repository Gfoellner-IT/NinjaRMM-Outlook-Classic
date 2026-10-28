using System;

namespace NinjaRmm.Outlook.Core
{
    public enum MatchAction { None, LogOnly, Propose, Move }

    // A match may refer to a closed ticket. Authorization is a separate, fresh check.
    public static class MatchingPolicy
    {
        public static MatchAction Decide(MatchResult result, MatchLevels levels, bool dryRun,
            TicketFacts currentTicket, bool rejected, bool mailCanMove)
        {
            if (result == null || levels == null || result.Outcome != MatchOutcome.Found) return MatchAction.None;
            MatchLevel level = levels.For(result.Rule);
            if (level == MatchLevel.Off) return MatchAction.None;
            if (dryRun || level == MatchLevel.LogOnly) return MatchAction.LogOnly;
            if (currentTicket == null || currentTicket.Id != result.TicketId || !currentTicket.IsOpen || !mailCanMove)
                return MatchAction.None;
            if (level == MatchLevel.Confirm) return rejected ? MatchAction.None : MatchAction.Propose;
            return MatchAction.Move;
        }

        public static bool CanConfirm(int proposedTicketId, TicketFacts refreshedTicket, bool mailCanMove)
        {
            return proposedTicketId > 0 && mailCanMove && refreshedTicket != null &&
                refreshedTicket.Id == proposedTicketId && refreshedTicket.IsOpen;
        }
    }

    public static class TicketTime
    {
        // Unknown dates remain null. The caller supplies the target time zone, including DST rules.
        public static DateTime? UtcToLocal(DateTime? utc, TimeZoneInfo zone)
        {
            if (zone == null) throw new ArgumentNullException("zone");
            if (!utc.HasValue) return null;
            if (utc.Value.Kind == DateTimeKind.Local) throw new ArgumentException("Expected UTC or unspecified UTC.", "utc");
            return TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(utc.Value, DateTimeKind.Utc), zone);
        }
    }
}
