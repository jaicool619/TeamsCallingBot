namespace TeamsCallingBot.Tda
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Newtonsoft.Json.Linq;
    using TeamsCallingBot.Config;

    /// <summary>
    /// Acquires and caches a client-credentials token scoped to the "TSL AI" resource, for talking to
    /// TDA (Tata Steel Digital Assistant). Deliberately mirrors the exact same pattern already proven
    /// in Chat/BotFrameworkChatClient.cs (token endpoint + client_id/client_secret/scope form POST,
    /// cache until near-expiry) - the only thing that differs between that class and this one is the
    /// token endpoint/scope (a different audience), not the mechanism.
    ///
    /// Inert by construction: <see cref="IsConfigured"/> is false until BaseUrl/TokenEndpoint/Scope are
    /// all set to real values (see TdaOptions - every field defaults empty). Callers must check
    /// IsConfigured before calling GetTokenAsync; this class never throws for "just not configured".
    /// </summary>
    public sealed class TdaTokenProvider
    {
        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        private readonly TdaOptions options;
        private readonly string clientId;
        private readonly string clientSecret;
        private readonly IGraphLogger logger;
        private readonly SemaphoreSlim tokenLock = new SemaphoreSlim(1, 1);

        private string cachedToken;
        private DateTime cachedTokenExpiresUtc = DateTime.MinValue;

        public TdaTokenProvider(TdaOptions options, string fallbackClientId, string fallbackClientSecret, IGraphLogger logger)
        {
            this.options = options ?? new TdaOptions();
            this.clientId = string.IsNullOrWhiteSpace(this.options.ClientId) ? fallbackClientId : this.options.ClientId;
            this.clientSecret = string.IsNullOrWhiteSpace(this.options.ClientSecret) ? fallbackClientSecret : this.options.ClientSecret;
            this.logger = logger;
        }

        /// <summary>True only once BaseUrl, TokenEndpoint, Scope, client id and secret are all real values.</summary>
        public bool IsConfigured =>
            this.options.Enabled &&
            !string.IsNullOrWhiteSpace(this.options.BaseUrl) &&
            !string.IsNullOrWhiteSpace(this.options.TokenEndpoint) &&
            !string.IsNullOrWhiteSpace(this.options.Scope) &&
            !string.IsNullOrWhiteSpace(this.clientId) &&
            !string.IsNullOrWhiteSpace(this.clientSecret);

        /// <summary>Returns a cached or freshly-acquired TSL AI scoped bearer token, or null if not configured / acquisition failed.</summary>
        public async Task<string> GetTokenAsync(CancellationToken cancellationToken = default)
        {
            if (!this.IsConfigured)
            {
                return null;
            }

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
                    new KeyValuePair<string, string>("client_id", this.clientId),
                    new KeyValuePair<string, string>("client_secret", this.clientSecret),
                    new KeyValuePair<string, string>("scope", this.options.Scope),
                });

                using (var response = await Http.PostAsync(this.options.TokenEndpoint, form, cancellationToken).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        this.logger?.Warn($"[TDA] TSL AI token request failed {(int)response.StatusCode}: {body}");
                        return null;
                    }

                    var json = JObject.Parse(body);
                    this.cachedToken = json.Value<string>("access_token");
                    int expiresIn = json.Value<int?>("expires_in") ?? 3600;
                    this.cachedTokenExpiresUtc = DateTime.UtcNow.AddSeconds(Math.Max(60, expiresIn - 120));
                    this.logger?.Info("[TDA] Acquired TSL AI scoped token.");
                    return this.cachedToken;
                }
            }
            catch (Exception ex)
            {
                this.logger?.Warn($"[TDA] TSL AI token acquisition threw: {ex.Message}");
                return null;
            }
            finally
            {
                this.tokenLock.Release();
            }
        }
    }
}
