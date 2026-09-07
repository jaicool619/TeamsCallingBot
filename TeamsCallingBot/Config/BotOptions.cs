namespace TeamsCallingBot.Config
{
    using System;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Plain configuration holder - deliberately NOT tied to Azure Cloud Services / RoleEnvironment
    /// (unlike Microsoft's original WorkerRole sample), because this runs as a plain process on a
    /// normal Azure VM, not a classic Cloud Service.
    /// </summary>
    public class BotOptions
    {
        public string AadAppId { get; set; }

        public string AadTenantId { get; set; }

        /// <summary>
        /// Either a client secret or a certificate thumbprint, depending on what Infra issues for app 5f6516f0.
        /// AuthenticationProvider currently only implements secret-based auth - see TODO there if a cert is issued instead.
        /// The SAME secret is used for the Bot Framework Connector token (chat posting) - see BotFrameworkChatClient.
        /// </summary>
        public string AadAppSecretOrCertThumbprint { get; set; }

        public Uri PlaceCallEndpointUrl { get; set; }

        /// <summary>
        /// A real DNS name (FQDN) is still needed here eventually - MediaPlatformInstanceSettings.ServiceFqdn
        /// and the certificate's subject/SAN are both matched against this for the media handshake.
        /// Using the bare IP as a stand-in for now (see InstancePublicIpAddress) is a temporary
        /// placeholder, not a real fix - a public cert generally can't be issued for a bare IP.
        /// </summary>
        public string ServiceDnsName { get; set; }

        /// <summary>
        /// The VM's real, reserved external IP (from the cloud console VM details page) - distinct from
        /// ServiceDnsName because MediaPlatformInstanceSettings needs both a literal IPAddress AND a separate FQDN.
        /// </summary>
        public string InstancePublicIpAddress { get; set; }

        public int InstancePublicPort { get; set; }

        public int MediaInstanceInternalPort { get; set; }

        public int MediaInstancePublicPort { get; set; }

        public string CertificateThumbprint { get; set; }

        /// <summary>
        /// TEMPORARY bypass for Conditional Access blocking this app's own MSAL client-credentials
        /// acquisition (see AuthenticationProvider's class comment for the full story). Paste a
        /// freshly-issued bearer token for app AadAppId here to skip MSAL entirely; leave blank/empty
        /// to use normal MSAL client-secret auth. The pasted token expires (~65 min observed) - refresh
        /// it here whenever calls start failing with 401s again.
        /// </summary>
        public string OverrideBearerToken { get; set; }

        /// <summary>
        /// Caps how many meetings this VM instance will join at once. Default of 3 matches the
        /// architecture doc's own "~3-4 concurrent meetings per capture VM" estimate for per-speaker
        /// (unmixed) audio - that number is explicitly marked "needs validation" there, so revisit
        /// once real throughput is measured on the actual VM hardware.
        /// </summary>
        public int MaxConcurrentCalls { get; set; } = 3;

        /// <summary>
        /// Where recordings/transcripts are written. Point this at wherever Google Drive for Desktop's
        /// synced folder lives on the VM to get everything into Drive automatically - defaults to a
        /// plain local folder if left blank (see RecordingsManager.ResolveDefaultDirectory).
        /// </summary>
        public string TranscriptOutputFolder { get; set; }

        // -------------------------------------------------------------------------------------
        // Video recording (screen share + participant camera)
        // -------------------------------------------------------------------------------------

        /// <summary>Record each participant's screen share (VBSS) to a per-person video file.</summary>
        public bool RecordScreenShare { get; set; } = true;

        /// <summary>Record the subscribed participant camera stream to a per-person video file.</summary>
        public bool RecordParticipantVideo { get; set; } = true;

        /// <summary>
        /// Target frames per second written to the screen-share video file. Teams delivers screen
        /// share at a variable rate (often 1-30 fps); frames are sampled/duplicated onto this fixed
        /// timeline so playback timing stays real-time. 5 fps is plenty for slides/documents and
        /// keeps CPU + disk low (1080p MJPEG at 5 fps is roughly 30-45 MB/minute).
        /// </summary>
        public int ScreenShareRecordingFps { get; set; } = 5;

        /// <summary>Target frames per second for participant camera recordings.</summary>
        public int ParticipantVideoRecordingFps { get; set; } = 3;

        /// <summary>JPEG quality (1-100) for recorded video frames and snapshots.</summary>
        public int VideoJpegQuality { get; set; } = 75;

        /// <summary>How often (seconds) a still JPEG snapshot of the shared screen is saved for the MoM.</summary>
        public int SnapshotIntervalSeconds { get; set; } = 10;

        /// <summary>
        /// A new video segment file is started once the current one exceeds this size (bytes).
        /// Plain RIFF AVI files must stay under 2 GB, so the default rolls over well before that.
        /// </summary>
        public long MaxVideoSegmentBytes { get; set; } = 1_500_000_000;

        /// <summary>
        /// Optional full path to ffmpeg.exe. When set and present, every finished AVI segment is
        /// re-encoded to H.264 MP4 at call end (the AVI is kept). Leave blank to skip.
        /// </summary>
        public string FfmpegPath { get; set; } = string.Empty;

        // -------------------------------------------------------------------------------------
        // Voice / greetings
        // -------------------------------------------------------------------------------------

        /// <summary>Speak the greeting (GreetingText) into the meeting once audio send becomes active.</summary>
        public bool SpeakGreetingOnJoin { get; set; } = true;

        /// <summary>Say a short "Hi &lt;name&gt;" when a human participant joins after the bot.</summary>
        public bool GreetParticipantsByName { get; set; } = true;

        /// <summary>Speak + post to chat when a participant starts sharing their screen.</summary>
        public bool AnnounceScreenShare { get; set; } = true;

        public string GreetingText { get; set; } =
            "Hello everyone, I am the AI meeting assistant. I have joined to record this meeting and prepare the minutes.";

        /// <summary>
        /// Culture of the Windows TTS voice to prefer, e.g. "en-IN" for Indian English.
        /// </summary>
        public string TtsCulture { get; set; } = "en-IN";

        /// <summary>Exact installed voice name to force, e.g. "Microsoft Zira Desktop".</summary>
        public string TtsVoiceName { get; set; } = string.Empty;

        /// <summary>"Female" or "Male" hint used when several voices match TtsCulture.</summary>
        public string TtsVoiceGender { get; set; } = "Female";

        // -------------------------------------------------------------------------------------
        // Chat posting via the Bot Framework Connector
        // -------------------------------------------------------------------------------------

        public bool UseBotFrameworkForChat { get; set; } = true;

        public string BotFrameworkServiceUrl { get; set; } = "https://smba.trafficmanager.net/teams/";

        // -------------------------------------------------------------------------------------
        // Minutes of Meeting generation
        // -------------------------------------------------------------------------------------

        public MomOptions Mom { get; set; } = new MomOptions();

        // -------------------------------------------------------------------------------------
        // Google Cloud Storage
        // -------------------------------------------------------------------------------------

        public GcsOptions Gcs { get; set; } = new GcsOptions();

        // -------------------------------------------------------------------------------------
        // Cloud Run relay
        // -------------------------------------------------------------------------------------

        public CloudRunRelayOptions CloudRunRelay { get; set; } = new CloudRunRelayOptions();

        /// <summary>
        /// Set once at startup (Startup.cs) so static classes without DI access (TranscriptSaver)
        /// can still read config. Deliberately simple - this process only ever loads one BotOptions.
        /// </summary>
        public static BotOptions Current { get; private set; }

        public static BotOptions LoadFromConfig()
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();

            var options = new BotOptions();
            config.GetSection("Bot").Bind(options);

            if (options.Mom == null)
            {
                options.Mom = new MomOptions();
            }

            if (options.Gcs == null)
            {
                options.Gcs = new GcsOptions();
            }

            if (options.CloudRunRelay == null)
            {
                options.CloudRunRelay = new CloudRunRelayOptions();
            }

            Current = options;
            return options;
        }

        public static BotOptions Reload()
        {
            return LoadFromConfig();
        }
    }

    public class MomOptions
    {
        public bool Enabled { get; set; } = true;
        public bool GenerateWordDocument { get; set; } = true;
        public string AnthropicApiKey { get; set; } = string.Empty;
        public string Model { get; set; } = "claude-opus-5";
        public int MaxSnapshotsForAi { get; set; } = 12;
        public string Language { get; set; } = "English";
        public string OrganizationName { get; set; } = string.Empty;
        public bool EmbedSnapshots { get; set; } = true;
        public int MaxSnapshotsInDocument { get; set; } = 20;
        public int RequestTimeoutSeconds { get; set; } = 600;
    }

    public class GcsOptions
    {
        public bool Enabled { get; set; } = false;
        public string BucketName { get; set; } = string.Empty;
        public string ServiceAccountKeyPath { get; set; } = string.Empty;
        public string ObjectPathTemplate { get; set; } = "meetings/{ChatThreadId}/{CallId}/08_minutes_of_meeting.docx";
        public string TranscriptTextObjectPathTemplate { get; set; } = "meetings/{ChatThreadId}/{CallId}/04_transcript.txt";
        public string TranscriptJsonObjectPathTemplate { get; set; } = "meetings/{ChatThreadId}/{CallId}/04_transcript.json";
        public int SignedUrlExpiryHours { get; set; } = 168; // 7 days
    }

    public class CloudRunRelayOptions
    {
        public bool Enabled { get; set; } = false;
        public string RelayUrl { get; set; } = string.Empty;
        public int RequestTimeoutSeconds { get; set; } = 30;
    }
}
