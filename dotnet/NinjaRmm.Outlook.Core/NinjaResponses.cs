using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace NinjaRmm.Outlook.Core
{
    public sealed class BoardPage
    {
        public IList<TicketFacts> Tickets { get; private set; }
        public int Cursor { get; private set; }
        public int RowCount { get; private set; }
        internal BoardPage(IList<TicketFacts> tickets, int cursor, int rows)
        { Tickets = tickets; Cursor = cursor; RowCount = rows; }
    }

    // Parses the VBA reference response shapes. Never logs response bodies or performs I/O.
    public static class NinjaResponses
    {
        private static JavaScriptSerializer Serializer()
        { return new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024, RecursionLimit = 64 }; }
        private static object Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try { return Serializer().DeserializeObject(json); }
            catch (ArgumentException) { return null; }
            catch (InvalidOperationException) { return null; }
        }
        private static object At(object node, params string[] path)
        {
            foreach (string key in path)
            {
                var map = node as Dictionary<string, object>;
                if (map == null || !map.TryGetValue(key, out node)) return null;
            }
            return node;
        }
        private static bool Number(object value, out double result)
        {
            result = 0;
            if (!(value is int || value is long || value is decimal || value is double)) return false;
            result = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
            return !double.IsNaN(result) && !double.IsInfinity(result);
        }
        public static TicketState ParseTicketState(int status, string body)
        {
            object ticket = Parse(body);
            if (status == 404)
            {
                string code = At(ticket, "resultCode") as string;
                return string.IsNullOrEmpty(code) || code == "TICKET_NOT_FOUND" ? TicketState.NotFound : TicketState.Unavailable;
            }
            if (status != 200 || !(ticket is Dictionary<string, object>)) return TicketState.Unavailable;
            if (object.Equals(true, At(ticket, "deleted"))) return TicketState.Closed;
            double id;
            if (!Number(At(ticket, "status", "statusId"), out id)) return TicketState.Unavailable;
            return Math.Round(id, MidpointRounding.ToEven) == 6000 ? TicketState.Closed : TicketState.Open;
        }
        public static string Subject(string body) { return At(Parse(body), "subject") as string ?? ""; }
        public static DateTime? ClosedDateUtc(string body)
        {
            var entries = Parse(body) as object[];
            if (entries == null) return null;
            double newest = 0;
            foreach (object entry in entries)
            {
                string status = At(entry, "changeDiff", "status", "new") as string;
                bool closed = string.Equals(status, "Closed", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(status, "Geschlossen", StringComparison.OrdinalIgnoreCase) || object.Equals("DELETE", At(entry, "type"));
                double time;
                if (closed && Number(At(entry, "createTime"), out time) && time > newest) newest = time;
            }
            if (newest <= 0) return null;
            try { return new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(Math.Truncate(newest)); }
            catch (ArgumentOutOfRangeException) { return null; }
        }
        public static BoardPage ParseBoard(string body)
        {
            object root = Parse(body);
            var rows = At(root, "data") as object[];
            if (rows == null) return null;
            double value; int cursor = 0;
            if (Number(At(root, "metadata", "lastCursorId"), out value) && value > 0 && value <= int.MaxValue && value == Math.Truncate(value)) cursor = (int)value;
            var tickets = new List<TicketFacts>();
            foreach (object row in rows)
            {
                if (!Number(At(row, "id"), out value) || value <= 0 || value > int.MaxValue || value != Math.Truncate(value)) continue;
                int id = (int)value;
                double state;
                bool open = Number(At(row, "status", "statusId"), out state) && Math.Round(state, MidpointRounding.ToEven) != 6000 &&
                    !object.Equals(true, At(row, "deleted"));
                tickets.Add(new TicketFacts(id, At(row, "summary") as string ?? "", open));
            }
            return new BoardPage(tickets.AsReadOnly(), cursor, rows.Length);
        }
        public static string BoardRequest(int pageSize, int cursor = 0, string summaryContains = null)
        {
            if (pageSize < 1 || pageSize > 1000) throw new ArgumentOutOfRangeException("pageSize");
            if (cursor < 0) throw new ArgumentOutOfRangeException("cursor");
            var body = new Dictionary<string, object> { { "pageSize", pageSize } };
            if (cursor > 0) body.Add("lastCursorId", cursor);
            if (!string.IsNullOrEmpty(summaryContains)) body.Add("filters", new[] {
                new { field = "summary", @operator = "contains", value = summaryContains } });
            return Serializer().Serialize(body);
        }
    }
}
