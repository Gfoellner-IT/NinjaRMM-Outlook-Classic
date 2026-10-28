using System;
using System.IO;
using NinjaRmm.Outlook.Core;

namespace NinjaRmm.Outlook.Core.Tests
{
    internal static partial class Program
    {
        private static string Fixture(string name)
        { return File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fixtures", name + ".json")); }
        private static void RunResponseTests()
        {
            Run("Api_ResponseStatesAndMalformedInput", ResponseStates);
            Run("Api_ClosingDatesFromSharedVbaFixtures", ResponseClosingDates);
            Run("Api_BoardPagesAndEscapedRequests", ResponseBoards);
        }
        private static void ResponseStates()
        {
            Equal(TicketState.Open, NinjaResponses.ParseTicketState(200, Fixture("ticket_open")));
            Equal(TicketState.Closed, NinjaResponses.ParseTicketState(200, Fixture("ticket_closed")));
            Equal(TicketState.Closed, NinjaResponses.ParseTicketState(200, Fixture("ticket_deleted")));
            Equal(TicketState.NotFound, NinjaResponses.ParseTicketState(404, Fixture("ticket_not_found")));
            Equal(TicketState.NotFound, NinjaResponses.ParseTicketState(404, ""));
            Equal(TicketState.Unavailable, NinjaResponses.ParseTicketState(404, "{\"resultCode\":\"BOARD_NOT_FOUND\"}"));
            foreach (int status in new[] { 0, 401, 403, 429, 500, 503 })
                Equal(TicketState.Unavailable, NinjaResponses.ParseTicketState(status, Fixture("ticket_closed")));
            foreach (string text in new[] { "", "not json", "[]", "null", "{}", "{\"status\":{\"statusId\":\"6000\"}}" })
                Equal(TicketState.Unavailable, NinjaResponses.ParseTicketState(200, text));
            Equal("Example open ticket", NinjaResponses.Subject(Fixture("ticket_open")));
            Equal("", NinjaResponses.Subject("{\"subject\":null}"));
        }
        private static void ResponseClosingDates()
        {
            var expected = new DateTime(2026, 10, 6, 14, 1, 28, DateTimeKind.Utc);
            Equal(expected, NinjaResponses.ClosedDateUtc(Fixture("log_closed_manual")).Value);
            Equal(expected, NinjaResponses.ClosedDateUtc(Fixture("log_closed_english")).Value);
            Equal(new DateTime(2026, 9, 1, 13, 55, 6, DateTimeKind.Utc), NinjaResponses.ClosedDateUtc(Fixture("log_closed_automation")).Value);
            Equal(new DateTime(2026, 10, 7, 19, 6, 40, DateTimeKind.Utc), NinjaResponses.ClosedDateUtc(Fixture("log_reopened")).Value);
            Equal(DateTimeKind.Utc, NinjaResponses.ClosedDateUtc(Fixture("log_deleted")).Value.Kind);
            Equal(false, NinjaResponses.ClosedDateUtc(Fixture("log_open_only")).HasValue);
            foreach (string text in new[] { "", "[]", "{}", "[null,1,\"x\",{}]", "not json" })
                Equal(false, NinjaResponses.ClosedDateUtc(text).HasValue);
        }
        private static void ResponseBoards()
        {
            var page = NinjaResponses.ParseBoard(Fixture("board_run_page1"));
            Equal(true, page != null); Equal(true, page.RowCount > 0); Equal(true, page.Cursor > 0);
            Equal(0, NinjaResponses.ParseBoard(Fixture("board_run_empty")).RowCount);
            Equal(true, NinjaResponses.ParseBoard("{}") == null);
            var malformed = NinjaResponses.ParseBoard("{\"data\":[{\"id\":1},{\"id\":2,\"status\":{\"statusId\":6000}},{\"id\":0},{\"id\":3,\"deleted\":true,\"status\":{\"statusId\":1000}}]}");
            Equal(3, malformed.Tickets.Count);
            foreach (var ticket in malformed.Tickets) Equal(false, ticket.IsOpen);
            string request = NinjaResponses.BoardRequest(200, 17, "quote\" slash\\ line\n");
            Equal(true, request.Contains("\"lastCursorId\":17"));
            Equal(true, request.Contains("\\\"")); Equal(true, request.Contains("\\n"));
            Throws<ArgumentOutOfRangeException>(() => NinjaResponses.BoardRequest(0));
            Throws<ArgumentOutOfRangeException>(() => NinjaResponses.BoardRequest(200, -1));
        }
    }
}
