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
        /// The VM's real, reserved external IP (e.g. from `gcloud compute instances describe` or the
        /// Cloud Console VM details page) - distinct from ServiceDnsName because MediaPlatformInstanceSettings
        /// needs both a literal IPAddress AND a separate FQDN.
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

        public string TestMeetingJoinUrl { get; set; }

        public System.Collections.Generic.List<string> TestMeetingJoinUrls { get; set; } = new System.Collections.Generic.List<string>();

        /// <summary>
        /// Where TranscriptSaver writes the two transcript files. Point this at wherever
        /// Google Drive for Desktop's synced folder lives on the VM (set up once, interactively,
        /// over RDP) to get transcripts into Drive automatically with zero API/auth code - defaults
        /// to a plain local folder if left blank.
        /// </summary>
        public string TranscriptOutputFolder { get; set; }

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
            Current = options;
            return options;
        }

        public static BotOptions Reload()
        {
            return LoadFromConfig();
        }
    }
}
