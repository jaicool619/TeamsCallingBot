namespace TeamsCallingBot.Bot
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph.Communications.Calls;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Client;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Graph.Communications.Resources;
    using Microsoft.Skype.Bots.Media;
    using TeamsCallingBot.Common;
    using TeamsCallingBot.Config;

    /// <summary>
    /// Builds the Graph Communications client and joins a specific meeting by its join URL, with a
    /// LOCAL (application-hosted) media session so we receive raw audio frames instead of Microsoft's
    /// service-hosted media.
    ///
    /// This combines two pieces verified separately from Microsoft's own sample repo, which never
    /// demonstrates them together in one sample:
    ///   - meeting-join-by-URL logic: Samples/V1.0Samples/RemoteMediaSamples/IncidentBot/Bot/Bot.cs
    ///   - local media session / AudioSocket setup: Samples/V1.0Samples/LocalMediaSamples/PolicyRecordingBot
    ///
    /// CONFIRMED (2026-09-03, via ildasm against the actual installed DLLs, not guessed): checked
    /// Microsoft.Graph.Communications.Calls.dll directly - JoinMeetingParameters has a real
    /// constructor `.ctor(ChatInfo, MeetingInfo, IMediaSession)`, and ILocalMediaSession (returned by
    /// CreateMediaSession below) implements IMediaSession. The join+local-media-session combination
    /// in JoinCallAsync is a real, valid call - this is no longer an open question.
    /// </summary>
    public class Bot : IDisposable
    {
        private readonly IGraphLogger graphLogger;
        private readonly SemaphoreSlim concurrentCallSlots;
        private readonly BotOptions options;

        public Bot(BotOptions options, IGraphLogger graphLogger)
        {
            this.graphLogger = graphLogger;
            this.options = options;

            // Cap per the architecture doc's own capacity estimate for per-speaker (unmixed) audio
            // capture: ~3-4 concurrent meetings per capture VM, vs ~15 for mixed-stream - per-speaker
            // costs ~5x the bandwidth/compute per meeting. Configurable because that estimate is
            // explicitly marked "needs validation" in the doc, not a measured number.
            this.concurrentCallSlots = new SemaphoreSlim(options.MaxConcurrentCalls);

            var name = this.GetType().Assembly.GetName().Name;
            var builder = new CommunicationsClientBuilder(name, options.AadAppId, graphLogger);

            var authProvider = new AuthenticationProvider(name, options.AadAppId, options.AadAppSecretOrCertThumbprint, graphLogger, options.OverrideBearerToken);

            // FIX (2026-09-04): fire-and-forget, started as early as possible (right here at Bot
            // construction, well before any real Graph notification can arrive) so the OpenID config
            // network fetch is already cached by the time it's actually needed - see WarmupAsync's own
            // comment for why this matters (measured 2.7-4.7s notification-ack latency, likely
            // dominated by this fetch on the first callback of every test run).
            _ = authProvider.WarmupAsync();

            var notificationUrl = new Uri($"https://{options.ServiceDnsName}:{options.InstancePublicPort}/api/calling/notification");

            builder.SetAuthenticationProvider(authProvider);
            builder.SetNotificationUrl(notificationUrl);
            builder.SetServiceBaseUrl(options.PlaceCallEndpointUrl);
            builder.SetMediaPlatformSettings(new MediaPlatformSettings
            {
                MediaPlatformInstanceSettings = new MediaPlatformInstanceSettings
                {
                    CertificateThumbprint = options.CertificateThumbprint,
                    InstanceInternalPort = options.MediaInstanceInternalPort,
                    InstancePublicIPAddress = System.Net.IPAddress.Parse(options.InstancePublicIpAddress),
                    InstancePublicPort = options.MediaInstancePublicPort,
                    ServiceFqdn = options.ServiceDnsName,
                },
                ApplicationId = options.AadAppId,
            });

            this.Client = builder.Build();
            this.Client.Calls().OnUpdated += this.CallsOnUpdated;
        }

        public ICommunicationsClient Client { get; }

        public ConcurrentDictionary<string, CallHandler> CallHandlers { get; } = new ConcurrentDictionary<string, CallHandler>();

        public void Dispose()
        {
            this.Client.Calls().OnUpdated -= this.CallsOnUpdated;
            this.Client?.Dispose();
            this.concurrentCallSlots.Dispose();
        }

        public async Task<ICall> JoinCallAsync(string meetingJoinUrl)
        {
            // Fails fast instead of silently overloading the VM past the capacity the architecture
            // doc itself calls out as the limiting factor for per-speaker capture. This is a real
            // decision point, not a nicety - past this limit, per-call audio/Whisper throughput
            // degrades for EVERY call already in progress, not just the new one.
            if (!this.concurrentCallSlots.Wait(0))
            {
                throw new InvalidOperationException(
                    $"Refusing to join - already at the configured concurrency cap ({this.CallHandlers.Count} active calls).");
            }

            try
            {
                var (chatInfo, meetingInfo, tenantId) = await JoinInfo.ParseJoinURLAsync(meetingJoinUrl).ConfigureAwait(false);

                var mediaSession = this.Client.CreateMediaSession(
                    new AudioSocketSettings
                    {
                        StreamDirections = StreamDirection.Sendrecv,
                        SupportedAudioFormat = AudioFormat.Pcm16K,
                        ReceiveUnmixedMeetingAudio = true,
                    },
                    new[]
                    {
                        new VideoSocketSettings
                        {
                            StreamDirections = StreamDirection.Sendrecv,
                            ReceiveColorFormat = VideoColorFormat.NV12,
                            // 15 fps formats: the status card is rendered in software, 15 fps halves the CPU
                            // cost vs 30 fps and Teams picks the best one (PreferredVideoSourceFormat) for the
                            // available bandwidth. CallHandler honours whichever Teams asks for.
                            SupportedSendVideoFormats = new List<VideoFormat>
                            {
                                VideoFormat.NV12_1280x720_15Fps,
                                VideoFormat.NV12_640x360_15Fps,
                            },
                        }
                    },
                    new VideoSocketSettings
                    {
                        StreamDirections = StreamDirection.Recvonly,
                        ReceiveColorFormat = VideoColorFormat.NV12,
                        MediaType = MediaType.Vbss,
                    },
                    null,
                    mediaSessionId: Guid.NewGuid());

                var scenarioId = Guid.NewGuid();

                // Confirmed real constructor - see class-level comment above.
                var joinParams = new JoinMeetingParameters(chatInfo, meetingInfo, mediaSession)
                {
                    // FIX (2026-09-04) for Graph error 7505 "Request authorization tenant mismatch":
                    // this was never set, so the SDK never attached an "X-Microsoft-Tenant" hint to the
                    // outbound request at all. See JoinInfo.ParseJoinURL's comment for the full traced
                    // plumbing (confirmed via ildasm, not guessed) of exactly where this value goes.
                    TenantId = tenantId,
                };

                var call = await this.Client.Calls().AddAsync(joinParams, scenarioId).ConfigureAwait(false);

                CallHandler handler;
                try
                {
                    handler = new CallHandler(call, this.graphLogger, chatInfo?.ThreadId, this.options?.OverrideBearerToken);
                }
                catch (Exception ex)
                {
                    // FIX (2026-09-04): CallHandler's constructor (heartbeat/OnUpdated/AudioSocket
                    // wiring) can throw for reasons independent of the join itself (e.g.
                    // GetLocalMediaSession()/AudioSocket not synchronously populated yet) - AddAsync
                    // above has ALREADY succeeded, so Graph now has a live call with ZERO tracking on
                    // this side: no heartbeat, no OnUpdated subscription, no wind-down ever wired. Left
                    // alone, that call sits orphaned on Graph's side, indistinguishable in logs from
                    // "Graph silently ended it" when actually this process abandoned it. Confirmed real
                    // via ildasm: ICall.DeleteAsync(bool handleHttpNotFoundInternally = false, ...)
                    // exists - hang the call up explicitly instead of leaving it dangling. Best-effort:
                    // the delete itself could also fail; that's logged and swallowed since the ORIGINAL
                    // exception, not this cleanup attempt, is what should actually surface to the caller.
                    this.graphLogger.Error(ex, $"CallHandler construction failed for call {call.Id} - hanging up the now-orphaned Graph-side call.");
                    try
                    {
                        await call.DeleteAsync().ConfigureAwait(false);
                    }
                    catch (Exception deleteEx)
                    {
                        this.graphLogger.Error(deleteEx, $"Failed to hang up orphaned call {call.Id} after CallHandler construction failed.");
                    }

                    throw;
                }

                this.CallHandlers[call.Id] = handler;

                this.graphLogger.Info($"Join call requested: {call.Id} ({this.CallHandlers.Count} active calls now).");
                return call;
            }
            catch
            {
                this.concurrentCallSlots.Release();
                throw;
            }
        }

        private void CallsOnUpdated(ICallCollection sender, CollectionEventArgs<ICall> args)
        {
            foreach (var call in args.RemovedResources)
            {
                if (this.CallHandlers.TryRemove(call.Id, out var handler))
                {
                    handler.Dispose();
                    this.concurrentCallSlots.Release();
                }
            }
        }
    }
}
