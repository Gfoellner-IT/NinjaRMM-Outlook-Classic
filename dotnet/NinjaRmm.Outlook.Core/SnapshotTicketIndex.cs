using System;
using System.Collections.Generic;
using System.Linq;

namespace NinjaRmm.Outlook.Core
{
    // Immutable ticket snapshot plus session-only conversation assignments, as in VBA.
    // Refresh/paging belong to infrastructure; this class cannot query an API or user files.
    public sealed class SnapshotTicketIndex : ITicketIndex
    {
        private readonly TicketFacts[] tickets;
        private readonly Dictionary<string, int> conversations = new Dictionary<string, int>(StringComparer.Ordinal);
        public SnapshotTicketIndex(IEnumerable<TicketFacts> tickets)
        {
            if (tickets == null) throw new ArgumentNullException("tickets");
            this.tickets = tickets.Where(t => t != null).ToArray();
            if (this.tickets.Select(t => t.Id).Distinct().Count() != this.tickets.Length)
                throw new ArgumentException("Duplicate ticket identifiers in snapshot.", "tickets");
        }
        public TicketFacts GetTicket(int id) { return tickets.FirstOrDefault(t => t.Id == id); }
        public IEnumerable<TicketFacts> FindByExternalNumber(string number)
        {
            if (string.IsNullOrEmpty(number)) return new TicketFacts[0];
            return tickets.Where(t => t.Summary.IndexOf(number, StringComparison.Ordinal) >= 0);
        }
        // Deliberately loose candidate selection; the matcher rechecks normalized titles/stems.
        public IEnumerable<TicketFacts> FindOpenBySubject(string subject) { return tickets.Where(t => t.IsOpen); }
        public IEnumerable<TicketFacts> FindOpenByStem(string stem) { return tickets.Where(t => t.IsOpen); }
        public int FindByConversation(string id)
        {
            int result;
            return id != null && conversations.TryGetValue(id, out result) ? result : 0;
        }
        public void RememberConversation(string id, int ticketId)
        {
            if (string.IsNullOrEmpty(id)) return;
            if (ticketId <= 0) throw new ArgumentOutOfRangeException("ticketId");
            conversations[id] = ticketId;
        }
    }
}
