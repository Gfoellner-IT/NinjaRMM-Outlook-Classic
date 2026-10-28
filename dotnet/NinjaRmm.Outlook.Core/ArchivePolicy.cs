using System;

namespace NinjaRmm.Outlook.Core
{
    public enum TicketState { Unavailable, Open, Closed, NotFound }
    public enum ArchiveDecision { Keep, ByApi, ByMailFallback }

    // Port of TicketDecision.bas. This class never accesses Outlook or the network.
    public static class ArchivePolicy
    {
        public static bool NeedsMailCheck(TicketState state, bool fallbackEnabled)
        {
            Validate(state);
            return state == TicketState.Unavailable && fallbackEnabled;
        }

        public static ArchiveDecision Decide(TicketState state, bool mailSaysClosed, bool fallbackEnabled)
        {
            Validate(state);
            if (state == TicketState.Closed || state == TicketState.NotFound)
                return ArchiveDecision.ByApi;
            if (state == TicketState.Open)
                return ArchiveDecision.Keep;
            return fallbackEnabled && mailSaysClosed
                ? ArchiveDecision.ByMailFallback : ArchiveDecision.Keep;
        }

        private static void Validate(TicketState state)
        {
            if (!Enum.IsDefined(typeof(TicketState), state))
                throw new ArgumentOutOfRangeException("state");
        }
    }
}
