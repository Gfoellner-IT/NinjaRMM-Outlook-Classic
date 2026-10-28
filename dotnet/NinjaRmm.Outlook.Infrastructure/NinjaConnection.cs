using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using NinjaRmm.Outlook.Core;

namespace NinjaRmm.Outlook.Infrastructure
{
    public interface ICredentialSource { ClientCredential Read(); }
    public sealed class ClientCredential
    {
        internal string Id { get; private set; }
        internal string Secret { get; private set; }
        public ClientCredential(string id, string secret)
        {
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrEmpty(secret))
                throw new ArgumentException("Client credentials are missing.");
            Id = id; Secret = secret;
        }
        public override string ToString() { return "[client credentials]"; }
    }
    public interface ITokenSource
    {
        Task<string> GetAsync(CancellationToken cancellation);
        void Invalidate();
    }
    public sealed class ClientCredentialsTokenSource : ITokenSource
    {
        private readonly HttpClient http;
        private readonly Uri endpoint;
        private readonly ICredentialSource credentials;
        private readonly Func<DateTime> utcNow;
        private readonly SemaphoreSlim gate = new SemaphoreSlim(1, 1);
        private string token;
        private DateTime expires;
        public ClientCredentialsTokenSource(HttpClient http, Uri endpoint, ICredentialSource credentials, Func<DateTime> utcNow = null)
        {
            if (http == null || credentials == null) throw new ArgumentNullException("http/credentials");
            ValidateHttps(endpoint);
            this.http = http; this.endpoint = endpoint; this.credentials = credentials;
            this.utcNow = utcNow ?? (() => DateTime.UtcNow);
        }
        public void Invalidate() { lock (gate) { token = null; expires = DateTime.MinValue; } }
        public async Task<string> GetAsync(CancellationToken cancellation)
        {
            await gate.WaitAsync(cancellation).ConfigureAwait(false);
            try
            {
                lock (gate) { if (!string.IsNullOrEmpty(token) && utcNow() < expires) return token; }
                var credential = credentials.Read();
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Content = new FormUrlEncodedContent(new Dictionary<string, string> {
                        { "grant_type", "client_credentials" }, { "client_id", credential.Id },
                        { "client_secret", credential.Secret }, { "scope", "monitoring" } });
                    using (var response = await http.SendAsync(request, cancellation).ConfigureAwait(false))
                    {
                        if (response.StatusCode != HttpStatusCode.OK) throw new InvalidOperationException("Token request failed (HTTP " + (int)response.StatusCode + ").");
                        Dictionary<string, object> parsed;
                        try { parsed = new JavaScriptSerializer().DeserializeObject(await response.Content.ReadAsStringAsync().ConfigureAwait(false)) as Dictionary<string, object>; }
                        catch (ArgumentException) { throw new InvalidOperationException("Invalid token response."); }
                        object access, lifetime;
                        if (parsed == null || !parsed.TryGetValue("access_token", out access) || !(access is string) || string.IsNullOrWhiteSpace((string)access) ||
                            !parsed.TryGetValue("expires_in", out lifetime) || !(lifetime is int) || (int)lifetime <= 0 || (int)lifetime > 86400)
                            throw new InvalidOperationException("Invalid token response.");
                        lock (gate)
                        {
                            token = (string)access;
                            expires = utcNow().AddSeconds(Math.Max(0, (int)lifetime - 60));
                            return token;
                        }
                    }
                }
            }
            finally { gate.Release(); }
        }
        internal static void ValidateHttps(Uri uri)
        {
            if (uri == null || !uri.IsAbsoluteUri || uri.Scheme != Uri.UriSchemeHttps ||
                uri.UserInfo.Length != 0 || uri.Query.Length != 0 || uri.Fragment.Length != 0)
                throw new ArgumentException("An absolute HTTPS endpoint without credentials, query or fragment is required.");
        }
    }

    // Public operations are limited to ticket/log reads and the read-only board-query POST.
    // A borrowed HttpClient/handler makes tests entirely synthetic. No requests run on construction.
    public sealed class ReadOnlyNinjaClient
    {
        private readonly HttpClient http;
        private readonly Uri apiBase;
        private readonly ITokenSource tokens;
        public ReadOnlyNinjaClient(HttpClient http, Uri apiBase, ITokenSource tokens)
        {
            ClientCredentialsTokenSource.ValidateHttps(apiBase);
            if (!apiBase.AbsolutePath.EndsWith("/", StringComparison.Ordinal)) throw new ArgumentException("API base must end with a slash.");
            if (http == null || tokens == null) throw new ArgumentNullException("http/tokens");
            this.http = http; this.apiBase = apiBase; this.tokens = tokens;
        }
        public static HttpClient CreateHttpClient()
        { return new HttpClient(new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(30) }; }

        private async Task<Tuple<int, string>> ReadAsync(string path, string json, CancellationToken cancellation)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
                using (var request = new HttpRequestMessage(json == null ? HttpMethod.Get : HttpMethod.Post, new Uri(apiBase, path)))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await tokens.GetAsync(cancellation).ConfigureAwait(false));
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    if (json != null) request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    using (var response = await http.SendAsync(request, cancellation).ConfigureAwait(false))
                    {
                        if (response.StatusCode == HttpStatusCode.Unauthorized && attempt == 0) { tokens.Invalidate(); continue; }
                        return Tuple.Create((int)response.StatusCode, await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                    }
                }
            }
            throw new InvalidOperationException("Read retry limit reached.");
        }
        public async Task<TicketFacts> GetTicketAsync(int id, CancellationToken cancellation)
        {
            if (id <= 0) throw new ArgumentOutOfRangeException("id");
            var response = await ReadAsync("ticketing/ticket/" + id.ToString(CultureInfo.InvariantCulture), null, cancellation).ConfigureAwait(false);
            TicketState state = NinjaResponses.ParseTicketState(response.Item1, response.Item2);
            if (state == TicketState.Unavailable || state == TicketState.NotFound) return null;
            return new TicketFacts(id, NinjaResponses.Subject(response.Item2), state == TicketState.Open);
        }
        public async Task<DateTime?> GetClosedDateUtcAsync(int id, CancellationToken cancellation)
        {
            if (id <= 0) throw new ArgumentOutOfRangeException("id");
            var response = await ReadAsync("ticketing/ticket/" + id.ToString(CultureInfo.InvariantCulture) + "/log-entry?type=SAVE", null, cancellation).ConfigureAwait(false);
            return response.Item1 == 200 ? NinjaResponses.ClosedDateUtc(response.Item2) : null;
        }
        public async Task<IList<TicketFacts>> QueryBoardAsync(int boardId, CancellationToken cancellation, string summaryContains = null)
        {
            if (boardId <= 0) throw new ArgumentOutOfRangeException("boardId");
            int cursor = 0;
            var cursors = new HashSet<int>(); var tickets = new Dictionary<int, TicketFacts>();
            for (int page = 0; page < 50; page++)
            {
                var response = await ReadAsync("ticketing/trigger/board/" + boardId.ToString(CultureInfo.InvariantCulture) + "/run",
                    NinjaResponses.BoardRequest(200, cursor, summaryContains), cancellation).ConfigureAwait(false);
                BoardPage parsed = response.Item1 == 200 ? NinjaResponses.ParseBoard(response.Item2) : null;
                if (parsed == null) throw new InvalidOperationException("Board response unavailable or invalid.");
                foreach (var ticket in parsed.Tickets) tickets[ticket.Id] = ticket;
                if (parsed.RowCount == 0 || parsed.Cursor == 0) return new List<TicketFacts>(tickets.Values).AsReadOnly();
                if (!cursors.Add(parsed.Cursor)) throw new InvalidOperationException("Board cursor repeated; snapshot is incomplete.");
                cursor = parsed.Cursor;
            }
            throw new InvalidOperationException("Board page limit reached; snapshot is incomplete.");
        }
    }
}
