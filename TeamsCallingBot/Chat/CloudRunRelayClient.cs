namespace TeamsCallingBot.Chat
{
    using System;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using Newtonsoft.Json;
    using TeamsCallingBot.Config;
    using TeamsCallingBot.Mom;

    /// <summary>
    /// Fires exactly one HTTP POST per finished meeting to a Cloud Run relay endpoint: a short MoM
    /// summary plus the signed GCS link to the one Word document (see GcsUploader). The VM has no
    /// network path to Power Automate or to email - Cloud Run is the only component with that
    /// reach - so this client's only job is handing off the payload; Cloud Run decides what to do
    /// with it (call Power Automate, post to Teams chat, send the email, etc).
    /// </summary>
    public class CloudRunRelayClient
    {
        private readonly CloudRunRelayOptions options;
        private readonly Action<string> log;

        public bool IsConfigured =>
            this.options != null && this.options.Enabled && !string.IsNullOrWhiteSpace(this.options.RelayUrl);

        public CloudRunRelayClient(CloudRunRelayOptions options, Action<string> log = null)
        {
            this.options = options ?? new CloudRunRelayOptions();
            this.log = log ?? (_ => { });
        }

        public sealed class Payload
        {
            [JsonProperty("callId")]
            public string CallId { get; set; }

            [JsonProperty("chatThreadId")]
            public string ChatThreadId { get; set; }

            [JsonProperty("meetingTitle")]
            public string MeetingTitle { get; set; }

            [JsonProperty("meetingDate")]
            public string MeetingDate { get; set; }

            [JsonProperty("durationMinutes")]
            public double DurationMinutes { get; set; }

            [JsonProperty("attendees")]
            public string[] Attendees { get; set; }

            [JsonProperty("executiveSummary")]
            public string ExecutiveSummary { get; set; }

            [JsonProperty("decisionsCount")]
            public int DecisionsCount { get; set; }

            [JsonProperty("actionItemsCount")]
            public int ActionItemsCount { get; set; }

            [JsonProperty("momDocumentUrl")]
            public string MomDocumentUrl { get; set; }

            [JsonProperty("momDocumentUrlExpiresAt")]
            public string MomDocumentUrlExpiresAt { get; set; }

            [JsonProperty("generatedByAi")]
            public bool GeneratedByAi { get; set; }
        }

        /// <summary>
        /// Sends the single combined notification to Cloud Run. Never throws - a relay outage
        /// should never block or fail call wind-down; failures are logged only. The MoM file itself
        /// is already safely uploaded to GCS by this point, independent of whether this call succeeds.
        /// </summary>
        public async Task<bool> SendMomNotificationAsync(MomDocument doc, string chatThreadId, string signedMomUrl, int signedUrlExpiryHours, bool usedAi)
        {
            if (!this.IsConfigured)
            {
                this.log("[CloudRunRelay] Not configured (Bot:CloudRunRelay:Enabled/RelayUrl) - skipping notification.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(signedMomUrl))
            {
                this.log("[CloudRunRelay] No MoM document URL available - skipping notification (upload may have failed or GCS is not configured).");
                return false;
            }

            try
            {
                var payload = new Payload
                {
                    CallId = doc?.CallId,
                    ChatThreadId = chatThreadId,
                    MeetingTitle = doc?.Title,
                    MeetingDate = doc?.MeetingDate,
                    DurationMinutes = doc?.DurationMinutes ?? 0,
                    Attendees = (doc?.Attendees ?? new System.Collections.Generic.List<MomAttendee>()).ConvertAll(a => a.Name).ToArray(),
                    ExecutiveSummary = doc?.ExecutiveSummary,
                    DecisionsCount = doc?.Decisions?.Count ?? 0,
                    ActionItemsCount = doc?.ActionItems?.Count ?? 0,
                    MomDocumentUrl = signedMomUrl,
                    MomDocumentUrlExpiresAt = DateTime.UtcNow.AddHours(Math.Max(1, signedUrlExpiryHours)).ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    GeneratedByAi = usedAi,
                };

                string json = JsonConvert.SerializeObject(payload, Formatting.None);

                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(5, this.options.RequestTimeoutSeconds)) })
                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                {
                    var response = await client.PostAsync(this.options.RelayUrl, content).ConfigureAwait(false);
                    if (response.IsSuccessStatusCode)
                    {
                        this.log($"[CloudRunRelay] MoM notification relayed successfully ({response.StatusCode}).");
                        return true;
                    }

                    var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    this.log($"[CloudRunRelay] Relay endpoint returned {response.StatusCode}: {body}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                this.log($"[CloudRunRelay] Notification failed: {ex.Message}");
                return false;
            }
        }
    }
}
