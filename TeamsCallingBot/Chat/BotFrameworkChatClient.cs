namespace TeamsCallingBot.Chat
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Posts messages into a Teams meeting chat through the Bot Framework Connector REST API
    /// instead of the Microsoft Graph chat API.
    ///
    /// Why: the app registration behind this bot is also a real Azure Bot Service registration
    /// (the Teams app manifest's bots[].botId equals the AAD app id). Azure Bot Service lets a bot
    /// that is already part of a conversation post activities to it via
    ///     POST {serviceUrl}v3/conversations/{conversationId}/activities
    /// authenticated with a client-credentials token whose audience is https://api.botframework.com.
    /// This is a completely separate permission system from Graph (no roles/scp claims, no
    /// ChatMessage.Send.* application permission, no tenant-admin consent). Once the calling bot has
    /// joined the meeting it is a member of the meeting chat, so this is the standard way Teams bots
    /// send proactive messages.
    ///
    /// Token endpoint: https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token
    ///   grant_type=client_credentials, client_id={AppId}, client_secret={secret},
    ///   scope=https://api.botframework.com/.default
    /// (This is the documented endpoint for public-cloud bots - the "botframework.com" tenant is
    /// literal, NOT the customer's tenant id.)
    ///
    /// The conversationId for a meeting chat is the chat thread id parsed from the join URL,
    /// e.g. 19:meeting_...@thread.v2. The serviceUrl for Teams is regional; the global entry
    /// https://smba.trafficmanager.net/teams/ works for proactive messages, regional values are
    /// .../amer/, .../emea/, .../apac/, .../in/. It is configurable (Bot:BotFrameworkServiceUrl).
    /// </summary>
    public sealed class BotFrameworkChatClient : IDisposable
    {
        private const string TokenEndpoint = "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token";
        private const string Scope = "https://api.botframework.com/.default";

        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        private readonly string appId;
        private readonly string appSecret;
        private readonly string serviceUrl;
        private readonly IGraphLogger logger;
        private readonly SemaphoreSlim tokenLock = new SemaphoreSlim(1, 1);

        private string cachedToken;
        private DateTime cachedTokenExpiresUtc = DateTime.MinValue;

        public BotFrameworkChatClient(string appId, string appSecret, string serviceUrl, IGraphLogger logger)
        {
            this.appId = appId;
            this.appSecret = appSecret;
            this.serviceUrl = string.IsNullOrWhiteSpace(serviceUrl) ? "https://smba.trafficmanager.net/teams/" : serviceUrl.TrimEnd('/') + "/";
            this.logger = logger;
        }

        /// <summary>True when an app id and a real (non-placeholder) secret are available.</summary>
        public bool IsConfigured =>
            !string.IsNullOrWhiteSpace(this.appId) &&
            !string.IsNullOrWhiteSpace(this.appSecret) &&
            !this.appSecret.StartsWith("TODO", StringComparison.OrdinalIgnoreCase) &&
            !this.appSecret.StartsWith("YOUR_", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Sends a message activity to the conversation. <paramref name="content"/> may contain the
        /// simple HTML subset Teams accepts (b, i, u, br, a, ul/li, p) - it is sent with textFormat
        /// "xml" so Teams renders the tags. Returns true on 2xx.
        /// </summary>
        public async Task<bool> SendMessageAsync(string conversationId, string content, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(conversationId) || string.IsNullOrWhiteSpace(content))
            {
                return false;
            }

            if (!this.IsConfigured)
            {
                this.logger?.Warn("[BotFramework] Not configured (app id or client secret missing/placeholder) - skipping chat post.");
                return false;
            }

            try
            {
                string token = await this.GetTokenAsync(cancellationToken).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    return false;
                }

                string url = $"{this.serviceUrl}v3/conversations/{Uri.EscapeDataString(conversationId)}/activities";

                var activity = new Dictionary<string, object>
                {
                    ["type"] = "message",
                    ["textFormat"] = "xml",
                    ["text"] = content,
                    ["from"] = new Dictionary<string, object> { ["id"] = $"28:{this.appId}" },
                    ["conversation"] = new Dictionary<string, object> { ["id"] = conversationId },
                    ["channelId"] = "msteams",
                    ["serviceUrl"] = this.serviceUrl,
                };

                using (var request = new HttpRequestMessage(HttpMethod.Post, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = new StringContent(JsonConvert.SerializeObject(activity), Encoding.UTF8, "application/json");

                    using (var response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                    {
                        string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (response.IsSuccessStatusCode)
                        {
                            this.logger?.Info($"[BotFramework] Posted message to conversation {conversationId}: {body}");
                            Console.WriteLine($">>> [BotFramework] Chat message posted via Connector API.");
                            return true;
                        }

                        this.logger?.Warn($"[BotFramework] Connector returned {(int)response.StatusCode} {response.ReasonPhrase} for {url}: {body}");
                        Console.WriteLine($">>> [BotFramework] Connector returned {(int)response.StatusCode}: {body}");

                        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                        {
                            this.cachedToken = null; // force refresh next time
                        }

                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger?.Error(ex, "[BotFramework] Failed posting message via Connector API.");
                return false;
            }
        }

        private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
        {
            if (!string.IsNullOrEmpty(this.cachedToken) && DateTime.UtcNow < this.cachedTokenExpiresUtc)
            {
                return this.cachedToken;
            }

            await this.tokenLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                if (!string.IsNullOrEmpty(this.cachedToken) && DateTime.UtcNow < this.cachedTokenExpiresUtc)
                {
                    return this.cachedToken;
                }

                var form = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("grant_type", "client_credentials"),
                    new KeyValuePair<string, string>("client_id", this.appId),
                    new KeyValuePair<string, string>("client_secret", this.appSecret),
                    new KeyValuePair<string, string>("scope", Scope),
                });

                using (var response = await Http.PostAsync(TokenEndpoint, form, cancellationToken).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        this.logger?.Warn($"[BotFramework] Token request failed {(int)response.StatusCode}: {body}");
                        Console.WriteLine($">>> [BotFramework] Token request failed {(int)response.StatusCode}: {body}");
                        return null;
                    }

                    var json = JObject.Parse(body);
                    this.cachedToken = json.Value<string>("access_token");
                    int expiresIn = json.Value<int?>("expires_in") ?? 3600;
                    this.cachedTokenExpiresUtc = DateTime.UtcNow.AddSeconds(Math.Max(60, expiresIn - 120));
                    this.logger?.Info("[BotFramework] Acquired Connector token.");
                    return this.cachedToken;
                }
            }
            finally
            {
                this.tokenLock.Release();
            }
        }

        public void Dispose()
        {
            this.tokenLock.Dispose();
        }
    }
}
