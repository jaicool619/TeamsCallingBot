namespace TeamsCallingBot.Tda
{
    using System;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using TeamsCallingBot.Config;

    /// <summary>
    /// Talks to TDA (Tata Steel Digital Assistant): "ask a question and get a text answer" (so the
    /// bot can speak it, see CallHandler.SpeakTdaAnswerAsync) and "send it a message" using a token
    /// scoped to the "TSL AI" resource (see TdaTokenProvider).
    ///
    /// PLACEHOLDER CLIENT: the request/response shapes below (field names "query"/"answer",
    /// "message") are guesses in the absence of a real TDA API contract - see
    /// BOT_CAPABILITY_EXPECTATIONS.md section 7. Adjust AskAsync/SendMessageAsync's payload and
    /// response parsing once the real contract is known; the token-acquisition and
    /// enabled/configured gating around them will not need to change.
    ///
    /// Every public method here fails soft (returns null/false, logs a warning) - this integration
    /// must never be able to break the call recording/audio/chat pipeline it is bolted onto.
    /// </summary>
    public sealed class TdaClient
    {
        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        private readonly TdaOptions options;
        private readonly TdaTokenProvider tokenProvider;
        private readonly IGraphLogger logger;

        public TdaClient(TdaOptions options, TdaTokenProvider tokenProvider, IGraphLogger logger)
        {
            this.options = options ?? new TdaOptions();
            this.tokenProvider = tokenProvider;
            this.logger = logger;
        }

        /// <summary>True only when TDA is explicitly enabled and its token provider is fully configured.</summary>
        public bool IsConfigured => this.options.Enabled && (this.tokenProvider?.IsConfigured ?? false) && !string.IsNullOrWhiteSpace(this.options.BaseUrl);

        /// <summary>Asks TDA a question, returns its text answer (truncated to Tda:MaxSpokenAnswerChars for TTS), or null if unavailable.</summary>
        public async Task<string> AskAsync(string query)
        {
            if (!this.IsConfigured || string.IsNullOrWhiteSpace(query))
            {
                return null;
            }

            try
            {
                string token = await this.tokenProvider.GetTokenAsync().ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    return null;
                }

                string url = this.options.BaseUrl.TrimEnd('/') + this.options.QueryEndpointPath;
                var payload = new JObject { ["query"] = query };

                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json");

                    using (var response = await Http.SendAsync(request).ConfigureAwait(false))
                    {
                        string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (!response.IsSuccessStatusCode)
                        {
                            this.logger?.Warn($"[TDA] Query returned {(int)response.StatusCode}: {body}");
                            return null;
                        }

                        var json = JObject.Parse(body);
                        string answer = json.Value<string>("answer") ?? json.Value<string>("response") ?? json.Value<string>("text");
                        if (string.IsNullOrWhiteSpace(answer))
                        {
                            return null;
                        }

                        int max = Math.Max(50, this.options.MaxSpokenAnswerChars);
                        return answer.Length > max ? answer.Substring(0, max) + "..." : answer;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger?.Warn($"[TDA] AskAsync threw: {ex.Message}");
                return null;
            }
        }

        /// <summary>Sends a message to TDA using the TSL AI scoped token. Returns true on 2xx.</summary>
        public async Task<bool> SendMessageAsync(string message)
        {
            if (!this.IsConfigured || string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            try
            {
                string token = await this.tokenProvider.GetTokenAsync().ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    return false;
                }

                string url = this.options.BaseUrl.TrimEnd('/') + this.options.MessageEndpointPath;
                var payload = new JObject { ["message"] = message };

                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json");

                    using (var response = await Http.SendAsync(request).ConfigureAwait(false))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            this.logger?.Info("[TDA] Message sent.");
                            return true;
                        }

                        string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        this.logger?.Warn($"[TDA] SendMessage returned {(int)response.StatusCode}: {body}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger?.Warn($"[TDA] SendMessageAsync threw: {ex.Message}");
                return false;
            }
        }
    }
}
