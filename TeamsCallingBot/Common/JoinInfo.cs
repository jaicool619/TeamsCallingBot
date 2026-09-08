// Adapted, near-verbatim, from Microsoft's own sample repo:
// Samples/Common/Sample.Common/Meetings/JoinInfo.cs (microsoft-graph-comms-samples, MIT licensed).
namespace TeamsCallingBot.Common
{
    using System;
    using System.Net;
    using System.Net.Http;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using Microsoft.Graph;

    /// <summary>
    /// Parses a Teams meeting join URL into the pieces the Calling SDK needs. Handles:
    ///  - the long/old "meetup-join" format, which carries the thread/message id and a
    ///    "?context={tid,oid,...}" blob directly in the URL - parsed locally, no network call.
    ///    Also handles that same blob arriving wrapped in a "launcher" redirect URL
    ///    (teams.microsoft.com/dl/launcher/launcher.html?url=...) with one extra layer of percent
    ///    encoding - confirmed live 2026-09-04, see UnwrapLauncherUrl below.
    ///  - the short "teams.microsoft.com/meet/&lt;id&gt;?p=&lt;passcode&gt;" format, BEST EFFORT ONLY:
    ///    attempts to follow HTTP redirects in case Teams's server resolves it straight to a
    ///    long-format URL for some short-link variant. CONFIRMED LIVE (via curl) that it does NOT
    ///    work for at least one real short link tested tonight - that one 302s to the launcher page
    ///    itself (a 200 OK HTML/SPA page), still carrying only the same short /meet/&lt;id&gt;?p=...
    ///    with no context blob anywhere, which only resolves inside the Teams client's own JS. Kept
    ///    anyway since it's a correctly-bounded, harmless attempt (5-hop cap) that may work for other
    ///    short-link variants - just don't assume it solves every short link. Same class of problem
    ///    production's resolveShortUrl()/fallbackMeetingLinkFromBQ in PRODUCTION/src/teamsBot.js
    ///    already has to work around with a BigQuery fallback instead of pure HTTP resolution.
    /// NOT YET HANDLED: live-event links and tenant-branded custom domains, if those ever show up.
    /// </summary>
    public static class JoinInfo
    {
        private static readonly Regex LongFormatRegex = new Regex(
            "https://teams\\.microsoft\\.com.*/(?<thread>[^/]+)/(?<message>[^/]+)\\?context=(?<context>{.*})",
            RegexOptions.IgnoreCase);

        private static readonly Regex ShortFormatRegex = new Regex(
            "^https://teams\\.(microsoft|live)\\.com/meet/", RegexOptions.IgnoreCase);

        public static async Task<(ChatInfo ChatInfo, MeetingInfo MeetingInfo, string TenantId)> ParseJoinURLAsync(string joinURL)
        {
            if (TryParseLongFormat(joinURL, out var result))
            {
                return result;
            }

            if (ShortFormatRegex.IsMatch(WebUtility.UrlDecode(joinURL).Split('?')[0]))
            {
                var resolvedUrl = await ResolveShortUrlAsync(joinURL).ConfigureAwait(false);
                if (TryParseLongFormat(resolvedUrl, out result))
                {
                    return result;
                }
            }

            throw new ArgumentException(
                $"Join URL cannot be parsed (long-format and short-link resolution both failed): {joinURL}.", nameof(joinURL));
        }

        /// <summary>
        /// ROBUSTNESS (2026-09-04): tonight alone we've seen 2 different real-world wrapper shapes
        /// around the same underlying 19:meeting_...@thread.v2 + ?context={tid,oid} link (a plain
        /// link, and a "launcher" redirect link needing UnwrapLauncherUrl below) - and there's no
        /// guarantee a third UI surface won't wrap it differently again. As long as a real
        /// 19:meeting_...@thread.v2 + context pair exists SOMEWHERE in the given string, however many
        /// extra times it's been percent-encoded by whatever produced it, this keeps decoding and
        /// re-checking until it either matches or genuinely can't be decoded any further.
        /// </summary>
        private static bool TryParseLongFormat(string joinURL, out (ChatInfo ChatInfo, MeetingInfo MeetingInfo, string TenantId) result)
        {
            var candidate = UnwrapLauncherUrl(joinURL);
            for (var decodePass = 0; decodePass < 5; decodePass++)
            {
                var match = LongFormatRegex.Match(candidate);
                if (match.Success)
                {
                    result = BuildResult(match);
                    return true;
                }

                var nextCandidate = WebUtility.UrlDecode(candidate);
                if (string.Equals(nextCandidate, candidate, StringComparison.Ordinal))
                {
                    break; // no further decoding changed anything - stop instead of looping forever
                }

                candidate = nextCandidate;
            }

            result = default;
            return false;
        }

        private static (ChatInfo ChatInfo, MeetingInfo MeetingInfo, string TenantId) BuildResult(Match match)
        {
            // FIX (2026-09-04): was DataContractJsonSerializer against a Context class with
            // Tid/Oid/MessageId members - that matches JSON property names CASE-SENSITIVELY, and
            // tonight alone we've seen both lowercase ("tid"/"oid", the original test link) and
            // capitalized ("Tid"/"Oid", the launcher-wrapped link) context JSON from real Teams links.
            // Whichever casing didn't match would silently deserialize to null instead of throwing - a
            // null Organizer.User.Id / tenant that could fail far downstream in a confusing way, not at
            // the parse step where it'd be obvious. This context JSON is always a simple flat
            // {"tid":"...","oid":"...","messageId":"..."} object (no nesting, no arrays) - not worth a
            // JSON library dependency (deliberately not using Newtonsoft or any other here), so this
            // extracts each value with a case-insensitive regex directly instead.
            var contextJson = match.Groups["context"].Value;
            var tid = ExtractJsonStringValue(contextJson, "tid");
            var oid = ExtractJsonStringValue(contextJson, "oid");
            var messageId = ExtractJsonStringValue(contextJson, "messageId");

            var chatInfo = new ChatInfo
            {
                ThreadId = match.Groups["thread"].Value,
                MessageId = match.Groups["message"].Value,
                ReplyChainMessageId = messageId,
            };

            var meetingInfo = new OrganizerMeetingInfo
            {
                Organizer = new IdentitySet
                {
                    User = new Identity { Id = oid },
                },
            };
            meetingInfo.Organizer.User.SetTenantId(tid);

            // CONFIRMED (2026-09-04, via ildasm against Microsoft.Graph.Communications.Calls.dll):
            // JoinMeetingParameters.TenantId is a real, public, settable property. Traced its actual
            // effect: CallCollectionExtensions.AddAsync copies parameters.TenantId onto the built
            // Call.TenantId, which StatefulCallCollection.GetGraphClient then uses to set
            // GraphClientContext.TenantId - and ONLY if that's non-null/non-whitespace does the SDK
            // attach an "X-Microsoft-Tenant" property to the outbound request, which GraphAuthClient
            // reads and passes as the `tenant` argument into AuthenticationProvider's
            // AuthenticateOutboundRequestAsync. Bot.cs's JoinMeetingParameters construction never set
            // this before tonight, so no tenant hint was ever sent with the request at all - the
            // likely real cause of Graph error 7505 "Request authorization tenant mismatch". Returning
            // it here (the MEETING's own tenant from the join URL, not a hardcoded app-level tenant) so
            // Bot.cs can set JoinMeetingParameters.TenantId with it - stays correct even for a bot
            // joining meetings across different tenants, not just this single-tenant test.
            return (chatInfo, meetingInfo, tid);
        }

        private static string ExtractJsonStringValue(string json, string propertyName)
        {
            var propMatch = Regex.Match(json, $"\"{propertyName}\"\\s*:\\s*\"(?<value>[^\"]*)\"", RegexOptions.IgnoreCase);
            return propMatch.Success ? propMatch.Groups["value"].Value : null;
        }

        /// <summary>
        /// FIX (2026-09-04): Teams sometimes hands out a "launcher" redirect link instead of a plain
        /// meetup-join link - e.g.
        /// https://teams.microsoft.com/dl/launcher/launcher.html?url=&lt;encoded-inner-link&gt;&amp;type=meetup-join&amp;...
        /// - confirmed live: a real link pasted into TestMeetingJoinUrl came in exactly this shape and
        /// threw "cannot be parsed" against the plain regex (no literal '{' after "context=", only
        /// doubly-encoded text). Traced the encoding by hand: the "url" query parameter's value is the
        /// real meetup-join link PLUS its own already-encoded ?context={...} parameter, with ONE EXTRA
        /// layer of percent-encoding applied on top of the whole thing (so the inner context's
        /// pre-existing %-signs become %25). This extracts the raw "url" parameter, decodes it ONCE
        /// here, and reconstructs a normal-looking https://teams.microsoft.com/... string -
        /// TryParseLongFormat's own decode loop then resolves the remaining layer exactly as it
        /// already does for a plain (non-launcher) link. Returns the input unchanged if it's not
        /// actually a launcher URL (no "url=" query parameter found) - zero effect on the
        /// already-working plain-link path.
        /// </summary>
        private static string UnwrapLauncherUrl(string joinURL)
        {
            // Support browser URLs from Teams Web (e.g. teams.microsoft.com/light-meetings/launch?coords=...)
            var coordsMatch = Regex.Match(joinURL, "[?&]coords=(?<coords>[^&]+)", RegexOptions.IgnoreCase);
            if (coordsMatch.Success)
            {
                try
                {
                    string rawBase64 = WebUtility.UrlDecode(coordsMatch.Groups["coords"].Value);
                    rawBase64 = rawBase64.PadRight(rawBase64.Length + (4 - rawBase64.Length % 4) % 4, '=');
                    byte[] bytes = Convert.FromBase64String(rawBase64);
                    string json = System.Text.Encoding.UTF8.GetString(bytes);
                    string meetingUrl = ExtractJsonStringValue(json, "meetingUrl");
                    if (!string.IsNullOrWhiteSpace(meetingUrl))
                    {
                        return meetingUrl;
                    }
                }
                catch { }
            }

            var match = Regex.Match(joinURL, "[?&]url=(?<inner>[^&]+)");
            if (!match.Success)
            {
                return joinURL;
            }

            var innerDecodedOnce = WebUtility.UrlDecode(match.Groups["inner"].Value);
            return innerDecodedOnce.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                ? innerDecodedOnce
                : "https://teams.microsoft.com" + innerDecodedOnce;
        }

        /// <summary>
        /// Follows HTTP redirects on a short join link, hoping Teams's server resolves it straight to
        /// a long-format URL. CONFIRMED LIVE (2026-09-04, via curl) that at least one real short link
        /// does NOT resolve this way - it 302s to the launcher page itself (a 200 OK HTML/SPA page)
        /// whose wrapped inner content is STILL the same short /meet/&lt;id&gt;?p=&lt;passcode&gt;
        /// format, no context blob anywhere. That specific numeric-id+passcode resolution only
        /// happens inside the Teams client's own JS (an authenticated call), which a bare HttpClient
        /// hop-follow cannot replicate. Kept as a best-effort attempt (5-hop cap, correctly bounded)
        /// in case some OTHER short-link variant genuinely does 302 straight to a real long-format
        /// URL - just don't assume this alone solves the short-link problem in general.
        /// </summary>
        private static async Task<string> ResolveShortUrlAsync(string shortUrl)
        {
            var currentUrl = shortUrl;
            using (var handler = new HttpClientHandler { AllowAutoRedirect = false })
            using (var client = new HttpClient(handler))
            {
                for (var hop = 0; hop < 5; hop++)
                {
                    using (var response = await client.GetAsync(currentUrl, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                    {
                        var isRedirect = (int)response.StatusCode >= 300 && (int)response.StatusCode < 400;
                        if (!isRedirect || response.Headers.Location == null)
                        {
                            return currentUrl;
                        }

                        var location = response.Headers.Location;
                        currentUrl = location.IsAbsoluteUri ? location.ToString() : new Uri(new Uri(currentUrl), location).ToString();
                    }
                }
            }

            return currentUrl;
        }
    }
}
