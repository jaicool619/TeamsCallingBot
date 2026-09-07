namespace TeamsCallingBot.Storage
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Google.Apis.Auth.OAuth2;
    using Google.Cloud.Storage.V1;
    using TeamsCallingBot.Config;

    /// <summary>
    /// Uploads finished session artefacts (the MoM Word document, and optionally the transcript
    /// files) to Google Cloud Storage and returns signed, time-limited download URLs. Deliberately
    /// narrow: this does NOT mirror the whole session folder (audio/video/snapshots stay local
    /// only) - only the specific files callers ask for. See CloudRunRelayClient for what happens
    /// with the returned URL.
    /// </summary>
    public class GcsUploader
    {
        private readonly GcsOptions options;
        private readonly Action<string> log;

        public bool IsConfigured =>
            this.options != null
            && this.options.Enabled
            && !string.IsNullOrWhiteSpace(this.options.BucketName)
            && !string.IsNullOrWhiteSpace(this.options.ServiceAccountKeyPath)
            && File.Exists(this.options.ServiceAccountKeyPath);

        public GcsUploader(GcsOptions options, Action<string> log = null)
        {
            this.options = options ?? new GcsOptions();
            this.log = log ?? (_ => { });
        }

        /// <summary>
        /// Uploads the MoM Word document using Gcs.ObjectPathTemplate and returns a signed URL.
        /// </summary>
        public Task<string> UploadMomDocumentAsync(string filePath, string callId, string chatThreadId)
        {
            string objectName = BuildObjectName(this.options.ObjectPathTemplate, callId, chatThreadId);
            return this.UploadFileAsync(
                filePath,
                objectName,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "MoM");
        }

        /// <summary>
        /// Uploads the human-readable and JSON transcript files (04_transcript.txt / .json) from a
        /// finished session folder, if present. Returns the signed URLs that were successfully
        /// produced (missing files are skipped, not an error - e.g. Whisper found nothing to say).
        /// </summary>
        public async Task<TranscriptUrls> UploadTranscriptFilesAsync(string sessionDirectory, string callId, string chatThreadId)
        {
            var result = new TranscriptUrls();

            string txtPath = Path.Combine(sessionDirectory, "04_transcript.txt");
            string jsonPath = Path.Combine(sessionDirectory, "04_transcript.json");

            string txtObject = BuildObjectName(this.options.TranscriptTextObjectPathTemplate, callId, chatThreadId);
            string jsonObject = BuildObjectName(this.options.TranscriptJsonObjectPathTemplate, callId, chatThreadId);

            result.TextUrl = await this.UploadFileAsync(txtPath, txtObject, "text/plain", "transcript (txt)").ConfigureAwait(false);
            result.JsonUrl = await this.UploadFileAsync(jsonPath, jsonObject, "application/json", "transcript (json)").ConfigureAwait(false);
            return result;
        }

        public sealed class TranscriptUrls
        {
            public string TextUrl { get; set; }

            public string JsonUrl { get; set; }
        }

        /// <summary>
        /// Uploads <paramref name="filePath"/> to gs://{BucketName}/{objectName} and returns a signed
        /// HTTPS URL valid for SignedUrlExpiryHours. Returns null (never throws) on any failure -
        /// missing file, GCS outage, bad credentials - so this never blocks call wind-down or the
        /// local copy, which is always saved first and is the source of truth regardless.
        /// </summary>
        public async Task<string> UploadFileAsync(string filePath, string objectName, string contentType, string label)
        {
            if (!this.IsConfigured)
            {
                this.log($"[GCS] Not configured (Bot:Gcs:Enabled/BucketName/ServiceAccountKeyPath) - skipping {label} upload.");
                return null;
            }

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                this.log($"[GCS] {label} file not found, nothing to upload: {filePath}");
                return null;
            }

            try
            {
                var credential = GoogleCredential.FromFile(this.options.ServiceAccountKeyPath);

                using (var client = await StorageClient.CreateAsync(credential).ConfigureAwait(false))
                using (var stream = File.OpenRead(filePath))
                {
                    await client.UploadObjectAsync(this.options.BucketName, objectName, contentType, stream).ConfigureAwait(false);
                }

                this.log($"[GCS] Uploaded {label} to gs://{this.options.BucketName}/{objectName}");

                var signer = UrlSigner.FromServiceAccountPath(this.options.ServiceAccountKeyPath);
                var expiry = TimeSpan.FromHours(Math.Max(1, this.options.SignedUrlExpiryHours));
                string signedUrl = await signer.SignAsync(this.options.BucketName, objectName, expiry).ConfigureAwait(false);

                this.log($"[GCS] Signed URL generated for {label} (expires in {expiry.TotalHours:0} hours).");
                return signedUrl;
            }
            catch (Exception ex)
            {
                this.log($"[GCS] {label} upload or signing failed (file stays local-only): {ex.Message}");
                return null;
            }
        }

        private static string BuildObjectName(string template, string callId, string chatThreadId)
        {
            string safeCallId = SanitizeForPath(callId);
            string safeThreadId = SanitizeForPath(string.IsNullOrWhiteSpace(chatThreadId) ? callId : chatThreadId);

            return template
                .Replace("{CallId}", safeCallId)
                .Replace("{ChatThreadId}", safeThreadId);
        }

        private static string SanitizeForPath(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "unknown";
            }

            var chars = value.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] == ':' || chars[i] == '\\' || chars[i] == '#' || chars[i] == '?' || chars[i] == '%')
                {
                    chars[i] = '_';
                }
            }

            return new string(chars);
        }
    }
}
