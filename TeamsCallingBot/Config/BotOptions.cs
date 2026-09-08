namespace TeamsCallingBot.Config
{
    using System;
    using System.Collections.Generic;
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
        /// <summary>
        /// Maximum number of concurrent calls this single bot instance handles.
        /// Configured for testing 5 concurrent meetings at the same time.
        /// </summary>
        public int MaxConcurrentCalls { get; set; } = 5;

        /// <summary>
        /// List of meeting join URLs to test joining multiple meetings simultaneously (up to MaxConcurrentCalls).
        /// </summary>
        public List<string> TestMeetingJoinUrls { get; set; } = new List<string>();

        /// <summary>
        /// Path to whisper-cli executable (e.g. C:\whisper\whisper-cli.exe).
        /// </summary>
        public string WhisperExecutablePath { get; set; } = @"C:\whisper\whisper-cli.exe";

        /// <summary>
        /// Path to the standard Whisper model (e.g. ggml-base.bin or ggml-small.bin).
        /// </summary>
        public string WhisperModelPath { get; set; } = @"C:\whisper\models\ggml-base.bin";

        /// <summary>
        /// Language mode for Whisper transcription ("auto" for automatic Hindi/English/Hinglish detection).
        /// </summary>
        public string WhisperLanguage { get; set; } = "auto";

        /// <summary>
        /// Prompt context for Whisper to guide vocabulary and enable accurate Hindi, English, and Hinglish transcription.
        /// </summary>
        public string WhisperPrompt { get; set; } = "Meeting conversation in English, Hindi, and Hinglish: Namaste, haan, theek hai, okay, right, discuss karte hain.";

        /// <summary>
        /// Beam size for Whisper decoding (5 = higher accuracy / lower hallucinations).
        /// </summary>
        public int WhisperBeamSize { get; set; } = 5;

        /// <summary>
        /// Master switch to start the bot in muted state (does not speak audio into the meeting,
        /// while still recording and transcribing via Whisper).
        /// </summary>
        public bool StartMuted { get; set; } = true;

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
        public bool SpeakGreetingOnJoin { get; set; } = false;

        /// <summary>Say a short "Hi &lt;name&gt;" when a human participant joins after the bot.</summary>
        public bool GreetParticipantsByName { get; set; } = false;

        /// <summary>Speak + post to chat when a participant starts sharing their screen.</summary>
        public bool AnnounceScreenShare { get; set; } = false;

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

        // -------------------------------------------------------------------------------------
        // Tata Steel Digital Assistant (TDA) integration
        // -------------------------------------------------------------------------------------

        public TdaOptions Tda { get; set; } = new TdaOptions();

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

            if (options.Tda == null)
            {
                options.Tda = new TdaOptions();
            }

            if (options.TestMeetingJoinUrls == null)
            {
                options.TestMeetingJoinUrls = new List<string>();
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

    /// <summary>
    /// Settings for talking to TDA (Tata Steel Digital Assistant): asking it a question so the bot
    /// can speak the answer into the meeting, and sending it messages via a client-credentials token
    /// scoped to the "TSL AI" resource. See Tda/TdaTokenProvider.cs and Tda/TdaClient.cs.
    ///
    /// DEFAULT-OFF BY DESIGN: every value below is a placeholder - the real base URL, scope and
    /// endpoint paths are not yet known. Enabled defaults to false so this integration stays completely
    /// inert - no token requests, no HTTP calls, no behaviour change - on any deployment until someone
    /// deliberately turns it on with real values. Every call site in CallHandler wraps TDA calls in
    /// try/catch so a misconfigured or unreachable TDA never affects audio/video recording or the
    /// existing chat/MoM pipeline.
    /// </summary>
    public class TdaOptions
    {
        /// <summary>Master on/off switch. Must be explicitly set true - stays off by default.</summary>
        public bool Enabled { get; set; } = false;

        /// <summary>TDA API base URL, e.g. https://tda.tatasteel.example.com/api. Placeholder - not yet known.</summary>
        public string BaseUrl { get; set; } = string.Empty;

        /// <summary>
        /// OAuth token endpoint used to acquire the "TSL AI" scoped token (client-credentials grant),
        /// e.g. https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token. Placeholder.
        /// </summary>
        public string TokenEndpoint { get; set; } = string.Empty;

        /// <summary>
        /// App (client) id used to request the TSL AI token. Leave blank to reuse Bot:AadAppId if the
        /// same app registration is authorized for the TDA/TSL AI scope.
        /// </summary>
        public string ClientId { get; set; } = string.Empty;

        /// <summary>
        /// Client secret for the above. Leave blank to reuse Bot:AadAppSecretOrCertThumbprint.
        /// Never commit a real value here - appsettings.json is .gitignored, same as the other secrets.
        /// </summary>
        public string ClientSecret { get; set; } = string.Empty;

        /// <summary>
        /// The exact "TSL AI" scope string to request, e.g. "api://{tsl-ai-app-id}/.default" or a
        /// named scope like "api://{tsl-ai-app-id}/Message.Send". Must come from whoever owns the
        /// TDA/TSL AI app registration - not guessable. Placeholder.
        /// </summary>
        public string Scope { get; set; } = string.Empty;

        /// <summary>Endpoint TDA exposes for "ask a question, get a text answer". Placeholder.</summary>
        public string QueryEndpointPath { get; set; } = "/query";

        /// <summary>Endpoint TDA exposes for "send it a message". Placeholder.</summary>
        public string MessageEndpointPath { get; set; } = "/message";

        /// <summary>HTTP timeout (seconds) for TDA calls.</summary>
        public int RequestTimeoutSeconds { get; set; } = 30;

        /// <summary>
        /// TDA answers are spoken via TTS - cap length so the bot doesn't read out a huge block of
        /// text into the meeting. Longer answers are truncated with "...".
        /// </summary>
        public int MaxSpokenAnswerChars { get; set; } = 600;
    }
}
