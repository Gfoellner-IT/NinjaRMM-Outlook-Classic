using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NinjaRmm.Outlook.Core;

namespace NinjaRmm.Outlook.Infrastructure.Tests
{
    internal static class Program
    {
        private static int groups, failures, assertions;
        private static readonly CancellationToken None = CancellationToken.None;
        private static int Main()
        {
            Run("Credentials_TokenCacheExpiryAndFormEncoding", TokenCache);
            Run("Credentials_InvalidResponsesAndRedactedErrors", TokenErrors);
            Run("Transport_One401RetryAndReadOnlyRoutes", Retry);
            Run("Transport_PagingAndRepeatedCursorGuard", Paging);
            Run("Transport_InvalidEndpointsAndCancellation", Boundaries);
            Run("Settings_Utf16IniAndNeutralDefaults", Settings);
            Run("Index_SnapshotIsolationAndConversationSession", Snapshot);
            Console.WriteLine("Infrastructure test groups passed: {0}, failed: {1}, assertions: {2}", groups - failures, failures, assertions);
            Console.WriteLine(failures == 0 ? "RESULT: OK" : "RESULT: FAILED");
            return failures == 0 ? 0 : 1;
        }
        private static void Run(string name, Action test)
        {
            groups++;
            try { test(); Console.WriteLine("PASS " + name); }
            catch (Exception ex) { failures++; Console.WriteLine("FAIL " + name + ": " + ex.GetType().Name); }
        }
        private static void Equal<T>(T expected, T actual)
        { assertions++; if (!object.Equals(expected, actual)) throw new InvalidOperationException("Assertion failed."); }
        private static void Throws<T>(Action test) where T : Exception
        { assertions++; try { test(); } catch (T) { return; } throw new InvalidOperationException("Expected exception."); }
        private sealed class Handler : HttpMessageHandler
        {
            internal Func<HttpRequestMessage, int, HttpResponseMessage> Respond;
            internal int Calls;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellation)
            { cancellation.ThrowIfCancellationRequested(); return Task.FromResult(Respond(request, ++Calls)); }
        }
        private sealed class Credentials : ICredentialSource
        {
            public ClientCredential Read() { return new ClientCredential("example+id", "example&secret"); }
        }
        private sealed class Tokens : ITokenSource
        {
            internal int Invalidations;
            public void Invalidate() { Invalidations++; }
            public Task<string> GetAsync(CancellationToken cancellation) { cancellation.ThrowIfCancellationRequested(); return Task.FromResult("synthetic-token"); }
        }
        private static HttpResponseMessage Response(int status, string body)
        { return new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(body, Encoding.UTF8, "application/json") }; }
        private static void TokenCache()
        {
            DateTime now = new DateTime(2026, 10, 28, 12, 0, 0, DateTimeKind.Utc);
            var handler = new Handler { Respond = (request, count) => {
                Equal(HttpMethod.Post, request.Method); Equal("https://auth.example.invalid/token", request.RequestUri.AbsoluteUri);
                string form = request.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                Equal(true, form.Contains("client_id=example%2Bid")); Equal(true, form.Contains("client_secret=example%26secret"));
                Equal(true, form.Contains("grant_type=client_credentials")); Equal(true, form.Contains("scope=monitoring"));
                return Response(200, "{\"access_token\":\"synthetic-token\",\"expires_in\":3600}");
            }};
            using (var http = new HttpClient(handler))
            {
                var tokens = new ClientCredentialsTokenSource(http, new Uri("https://auth.example.invalid/token"), new Credentials(), () => now);
                Equal("synthetic-token", tokens.GetAsync(None).GetAwaiter().GetResult());
                Equal("synthetic-token", tokens.GetAsync(None).GetAwaiter().GetResult()); Equal(1, handler.Calls);
                now = now.AddMinutes(59); tokens.GetAsync(None).GetAwaiter().GetResult(); Equal(2, handler.Calls);
                tokens.Invalidate(); tokens.GetAsync(None).GetAwaiter().GetResult(); Equal(3, handler.Calls);
            }
            Equal("[client credentials]", new Credentials().Read().ToString());
        }
        private static void TokenErrors()
        {
            foreach (string body in new[] { "not json", "{}", "{\"access_token\":\"x\",\"expires_in\":-1}", "{\"access_token\":\"x\",\"expires_in\":\"3600\"}" })
            {
                using (var http = new HttpClient(new Handler { Respond = (r, n) => Response(200, body) }))
                {
                    var source = new ClientCredentialsTokenSource(http, new Uri("https://auth.example.invalid/token"), new Credentials());
                    Throws<InvalidOperationException>(() => source.GetAsync(None).GetAwaiter().GetResult());
                }
            }
            using (var http = new HttpClient(new Handler { Respond = (r, n) => Response(401, "example&secret") }))
            {
                try { new ClientCredentialsTokenSource(http, new Uri("https://auth.example.invalid/token"), new Credentials()).GetAsync(None).GetAwaiter().GetResult(); }
                catch (InvalidOperationException ex) { Equal(false, ex.ToString().Contains("example&secret")); }
            }
        }
        private static void Retry()
        {
            var tokens = new Tokens();
            var handler = new Handler { Respond = (r, n) => {
                Equal(HttpMethod.Get, r.Method); Equal("/v2/ticketing/ticket/20108", r.RequestUri.AbsolutePath);
                Equal("Bearer", r.Headers.Authorization.Scheme); Equal("synthetic-token", r.Headers.Authorization.Parameter);
                return n == 1 ? Response(401, "") : Response(200, "{\"subject\":\"Synthetic request\",\"status\":{\"statusId\":1000}}");
            }};
            using (var http = new HttpClient(handler))
            {
                var client = new ReadOnlyNinjaClient(http, new Uri("https://tenant.example.invalid/v2/"), tokens);
                Equal(20108, client.GetTicketAsync(20108, None).GetAwaiter().GetResult().Id); Equal(2, handler.Calls); Equal(1, tokens.Invalidations);
            }
            handler = new Handler { Respond = (r, n) => Response(401, "") }; tokens = new Tokens();
            using (var http = new HttpClient(handler))
            {
                Equal(true, new ReadOnlyNinjaClient(http, new Uri("https://tenant.example.invalid/v2/"), tokens).GetTicketAsync(20108, None).GetAwaiter().GetResult() == null);
                Equal(2, handler.Calls); Equal(1, tokens.Invalidations);
            }
        }
        private static void Paging()
        {
            var handler = new Handler { Respond = (r, n) => {
                Equal(HttpMethod.Post, r.Method); Equal("/v2/ticketing/trigger/board/2/run", r.RequestUri.AbsolutePath);
                if (n == 2) Equal(true, r.Content.ReadAsStringAsync().GetAwaiter().GetResult().Contains("\"lastCursorId\":17"));
                return n == 1 ? Response(200, "{\"data\":[{\"id\":20108,\"status\":{\"statusId\":1000}}],\"metadata\":{\"lastCursorId\":17}}") : Response(200, "{\"data\":[]}");
            }};
            using (var http = new HttpClient(handler))
            { Equal(1, new ReadOnlyNinjaClient(http, new Uri("https://tenant.example.invalid/v2/"), new Tokens()).QueryBoardAsync(2, None).GetAwaiter().GetResult().Count); }
            handler = new Handler { Respond = (r, n) => Response(200, "{\"data\":[{\"id\":20108}],\"metadata\":{\"lastCursorId\":17}}") };
            using (var http = new HttpClient(handler))
            { Throws<InvalidOperationException>(() => new ReadOnlyNinjaClient(http, new Uri("https://tenant.example.invalid/v2/"), new Tokens()).QueryBoardAsync(2, None).GetAwaiter().GetResult()); Equal(2, handler.Calls); }
        }
        private static void Boundaries()
        {
            var handler = new Handler { Respond = (r, n) => { throw new HttpRequestException("Simulated offline"); } };
            using (var http = new HttpClient(handler))
            {
                foreach (string uri in new[] { "http://tenant.example.invalid/", "https://user:pass@tenant.example.invalid/", "https://tenant.example.invalid/?x=1", "https://tenant.example.invalid/#x", "https://tenant.example.invalid/v2" })
                    Throws<ArgumentException>(() => new ReadOnlyNinjaClient(http, new Uri(uri), new Tokens()));
                var client = new ReadOnlyNinjaClient(http, new Uri("https://tenant.example.invalid/v2/"), new Tokens());
                Throws<ArgumentOutOfRangeException>(() => client.GetTicketAsync(0, None).GetAwaiter().GetResult());
                Throws<HttpRequestException>(() => client.GetTicketAsync(20108, None).GetAwaiter().GetResult());
                using (var cancel = new CancellationTokenSource()) { cancel.Cancel(); Throws<OperationCanceledException>(() => client.GetTicketAsync(20108, cancel.Token).GetAwaiter().GetResult()); }
            }
        }
        private static void Settings()
        {
            string file = Path.Combine(Path.GetTempPath(), "ninja-m9-test-" + Guid.NewGuid().ToString("N") + ".ini");
            try
            {
                var settings = new LocalSettings(file);
                Equal("https://tenant.example.invalid/v2/", settings.ApiBase.AbsoluteUri); Equal(2, settings.AllTicketsBoard);
                File.WriteAllText(file, "[Settings]\r\nAllowedSender=Example B\u00fcro\r\nAllTicketsBoard=7\r\n", Encoding.Unicode);
                Equal("Example B\u00fcro", settings.Get("AllowedSender", "")); Equal(7, settings.AllTicketsBoard);
                File.WriteAllText(file, "[Settings]\r\nApiBaseUrl=http://example.invalid/\r\nAllTicketsBoard=0\r\n", Encoding.Unicode);
                Throws<ArgumentException>(() => { var ignored = settings.ApiBase; });
                Throws<InvalidOperationException>(() => { var ignored = settings.AllTicketsBoard; });
            }
            finally { if (File.Exists(file)) File.Delete(file); }
        }
        private static void Snapshot()
        {
            var source = new List<TicketFacts> { new TicketFacts(20108, "Example warehouse configuration request", true) };
            var index = new SnapshotTicketIndex(source); source.Clear(); Equal(true, index.GetTicket(20108) != null);
            index.RememberConversation("example", 20108); Equal(20108, index.FindByConversation("example"));
            Equal(0, new SnapshotTicketIndex(new TicketFacts[0]).FindByConversation("example"));
            Throws<ArgumentException>(() => new SnapshotTicketIndex(new[] { index.GetTicket(20108), index.GetTicket(20108) }));
        }
    }
}
