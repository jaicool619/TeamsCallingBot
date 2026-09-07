namespace TeamsCallingBot.Mom
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using TeamsCallingBot.Config;
    using TeamsCallingBot.Video;

    /// <summary>
    /// Turns the locally-built MoM draft into detailed, human-quality minutes by "hearing" (the
    /// speaker-attributed transcript) and "seeing" (screen-share snapshots sent as images) with
    /// Claude via the Anthropic Messages API (POST https://api.anthropic.com/v1/messages).
    ///
    /// Plain HTTPS + Newtonsoft is used instead of the official Anthropic C# SDK on purpose: the SDK
    /// only ships netstandard2.0/net8/net9 builds with modern System.Text.Json dependencies, which
    /// clash with this net472 + ASP.NET Core 2.2 media bot's binding redirects. The request shape
    /// below follows the current API: model claude-opus-5, adaptive thinking, structured JSON output
    /// via output_config.format, and server-side refusal fallbacks.
    ///
    /// Fully optional: only used when Bot:Mom:AnthropicApiKey (or ANTHROPIC_API_KEY) is set.
    /// Nothing is sent anywhere otherwise.
    /// </summary>
    public sealed class ClaudeMomSummarizer
    {
        private const string Endpoint = "https://api.anthropic.com/v1/messages";
        private const string ApiVersion = "2023-06-01";
        private const string FallbackBeta = "server-side-fallback-2026-07-01";

        private readonly MomOptions options;
        private readonly Action<string> log;
        private readonly string apiKey;

        public ClaudeMomSummarizer(MomOptions options, Action<string> log)
        {
            this.options = options ?? new MomOptions();
            this.log = log ?? (_ => { });
            this.apiKey = !string.IsNullOrWhiteSpace(this.options.AnthropicApiKey)
                ? this.options.AnthropicApiKey.Trim()
                : Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        }

        public bool IsConfigured => !string.IsNullOrWhiteSpace(this.apiKey);

        /// <summary>
        /// Returns a refined document, or the original draft if the request fails for any reason.
        /// </summary>
        public async Task<MomDocument> SummarizeAsync(MomDocument draft, IList<string> snapshotPaths, CancellationToken cancellationToken = default)
        {
            if (draft == null || !this.IsConfigured)
            {
                return draft;
            }

            try
            {
                var content = new JArray();

                // Images first (the API recommends images before the text that refers to them).
                var chosen = ChooseSnapshots(snapshotPaths, this.options.MaxSnapshotsForAi);
                int imageNo = 1;
                foreach (var path in chosen)
                {
                    string b64 = LoadDownscaledJpegBase64(path, 1280, 70);
                    if (b64 == null)
                    {
                        continue;
                    }

                    content.Add(new JObject
                    {
                        ["type"] = "text",
                        ["text"] = $"Screen-share snapshot {imageNo} - file {Path.GetFileName(path)} (presenter/time are encoded in the file name):",
                    });
                    content.Add(new JObject
                    {
                        ["type"] = "image",
                        ["source"] = new JObject { ["type"] = "base64", ["media_type"] = "image/jpeg", ["data"] = b64 },
                    });
                    imageNo++;
                }

                content.Add(new JObject { ["type"] = "text", ["text"] = BuildPrompt(draft, chosen.Count) });

                var request = new JObject
                {
                    ["model"] = string.IsNullOrWhiteSpace(this.options.Model) ? "claude-opus-5" : this.options.Model,
                    ["max_tokens"] = 16000,
                    ["thinking"] = new JObject { ["type"] = "adaptive" },
                    ["fallbacks"] = "default",
                    ["system"] = SystemPrompt(this.options.Language),
                    ["messages"] = new JArray(new JObject { ["role"] = "user", ["content"] = content }),
                    ["output_config"] = new JObject
                    {
                        ["format"] = new JObject { ["type"] = "json_schema", ["schema"] = OutputSchema() },
                    },
                };

                string responseText = await this.PostAsync(request, cancellationToken).ConfigureAwait(false);
                if (responseText == null)
                {
                    // Retry once without structured output in case the account/model rejects output_config.
                    request.Remove("output_config");
                    responseText = await this.PostAsync(request, cancellationToken).ConfigureAwait(false);
                }

                if (string.IsNullOrWhiteSpace(responseText))
                {
                    return draft;
                }

                var refined = ParseDocument(responseText);
                if (refined == null)
                {
                    this.log("[MoM/Claude] Response was not parseable JSON - keeping the local draft and attaching the narrative.");
                    draft.ExecutiveSummary = responseText.Trim();
                    draft.GeneratedBy = $"Claude ({request["model"]}) narrative + local structure";
                    return draft;
                }

                return Merge(draft, refined, request["model"].ToString());
            }
            catch (Exception ex)
            {
                this.log($"[MoM/Claude] Summarisation failed - using local draft. {ex.GetType().Name}: {ex.Message}");
                return draft;
            }
        }

        // ------------------------------------------------------------------

        private async Task<string> PostAsync(JObject request, CancellationToken cancellationToken)
        {
            using (var http = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(60, this.options.RequestTimeoutSeconds)) })
            using (var message = new HttpRequestMessage(HttpMethod.Post, Endpoint))
            {
                message.Headers.Add("x-api-key", this.apiKey);
                message.Headers.Add("anthropic-version", ApiVersion);
                message.Headers.Add("anthropic-beta", FallbackBeta);
                message.Content = new StringContent(request.ToString(Formatting.None), Encoding.UTF8, "application/json");

                this.log($"[MoM/Claude] Sending summarisation request to {request["model"]} ...");
                using (var response = await http.SendAsync(message, cancellationToken).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        this.log($"[MoM/Claude] API returned {(int)response.StatusCode}: {Truncate(body, 600)}");
                        return null;
                    }

                    var json = JObject.Parse(body);
                    string stopReason = json.Value<string>("stop_reason");
                    if (stopReason == "refusal")
                    {
                        var details = json["stop_details"];
                        this.log($"[MoM/Claude] Request was refused: {details}");
                        return null;
                    }

                    var usage = json["usage"];
                    if (usage != null)
                    {
                        this.log($"[MoM/Claude] Usage: input={usage["input_tokens"]}, output={usage["output_tokens"]}, stop_reason={stopReason}");
                    }

                    var sb = new StringBuilder();
                    foreach (var block in json["content"] as JArray ?? new JArray())
                    {
                        if (block.Value<string>("type") == "text")
                        {
                            sb.Append(block.Value<string>("text"));
                        }
                    }

                    return sb.ToString();
                }
            }
        }

        private static string SystemPrompt(string language)
        {
            return
                "You are an expert executive assistant who writes detailed, accurate minutes of meeting (MoM) for corporate meetings. " +
                "You receive (1) a speaker-attributed transcript produced by local speech recognition - it may contain recognition errors, " +
                "so infer intended meaning but never invent facts; (2) metadata about attendees and screen-sharing sessions; and (3) snapshots " +
                "of the screens that were shared. Use the snapshots to understand what was presented (slides, documents, dashboards, code, " +
                "spreadsheets), read visible text/numbers and connect them to the discussion. " +
                $"Write in {language}. Be specific: name owners, dates, figures and systems exactly as stated or shown. " +
                "Where the transcript is unclear, say so briefly rather than guessing. Return ONLY the JSON object requested.";
        }

        private static string BuildPrompt(MomDocument draft, int imageCount)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Produce the minutes of this meeting as a JSON object with the requested schema.");
            sb.AppendLine();
            sb.AppendLine("## Meeting metadata");
            sb.AppendLine($"- Date: {draft.MeetingDate}");
            sb.AppendLine($"- Time: {draft.StartTime} to {draft.EndTime} ({draft.DurationMinutes} minutes)");
            sb.AppendLine($"- Attendees: {string.Join("; ", draft.Attendees.Select(a => $"{a.Name} (joined {a.JoinedAt}, left {a.LeftAt}, spoke ~{a.TalkTimeMinutes} min)"))}");
            sb.AppendLine();
            sb.AppendLine("## Screen sharing sessions");
            if (draft.ScreenShareObservations.Count == 0)
            {
                sb.AppendLine("- None");
            }

            foreach (var s in draft.ScreenShareObservations)
            {
                sb.AppendLine($"- {s.Presenter} shared from {s.StartTime} to {s.EndTime} ({s.DurationMinutes} min). Snapshots: {string.Join(", ", s.SnapshotFiles)}");
            }

            sb.AppendLine($"({imageCount} snapshot image(s) are attached above.)");
            sb.AppendLine();
            sb.AppendLine("## Transcript (timestamp | speaker | text)");
            foreach (var line in draft.Transcript)
            {
                sb.AppendLine($"{line.Timestamp} | {line.Speaker} | {line.Text}");
            }

            sb.AppendLine();
            sb.AppendLine("## Instructions");
            sb.AppendLine("- executiveSummary: 1-3 paragraphs covering purpose, what was discussed, outcomes.");
            sb.AppendLine("- topics: group the discussion into logical agenda topics (not fixed time blocks), each with a detailed discussion narrative and highlights.");
            sb.AppendLine("- decisions: only genuine decisions/agreements.");
            sb.AppendLine("- actionItems: every commitment, with owner (person name if identifiable, else 'Unassigned'), dueDate ('Not specified' if none), status 'Open'.");
            sb.AppendLine("- screenShareObservations: for each sharing session describe what was actually shown based on the snapshots (titles, figures, tables, charts, code, errors) and how it related to the discussion.");
            sb.AppendLine("- openQuestions and nextSteps: as discussed.");
            sb.AppendLine("- attendees: keep the given names; add a role if it is evident (e.g. 'Presenter', 'Organizer').");
            return sb.ToString();
        }

        private static JObject OutputSchema()
        {
            JObject StringArray() => new JObject { ["type"] = "array", ["items"] = new JObject { ["type"] = "string" } };
            JObject Str() => new JObject { ["type"] = "string" };
            JObject Num() => new JObject { ["type"] = "number" };

            JObject Obj(JObject props)
            {
                return new JObject
                {
                    ["type"] = "object",
                    ["properties"] = props,
                    ["required"] = new JArray(props.Properties().Select(p => p.Name)),
                    ["additionalProperties"] = false,
                };
            }

            var attendee = Obj(new JObject { ["name"] = Str(), ["role"] = Str(), ["joinedAt"] = Str(), ["leftAt"] = Str(), ["talkTimeMinutes"] = Num() });
            var topic = Obj(new JObject { ["title"] = Str(), ["timeRange"] = Str(), ["discussion"] = Str(), ["highlights"] = StringArray() });
            var action = Obj(new JObject { ["task"] = Str(), ["owner"] = Str(), ["dueDate"] = Str(), ["status"] = Str(), ["source"] = Str() });
            var share = Obj(new JObject { ["presenter"] = Str(), ["startTime"] = Str(), ["endTime"] = Str(), ["durationMinutes"] = Num(), ["summary"] = Str(), ["contentObserved"] = StringArray() });

            return Obj(new JObject
            {
                ["title"] = Str(),
                ["executiveSummary"] = Str(),
                ["keyPoints"] = StringArray(),
                ["attendees"] = new JObject { ["type"] = "array", ["items"] = attendee },
                ["topics"] = new JObject { ["type"] = "array", ["items"] = topic },
                ["decisions"] = StringArray(),
                ["actionItems"] = new JObject { ["type"] = "array", ["items"] = action },
                ["openQuestions"] = StringArray(),
                ["nextSteps"] = StringArray(),
                ["screenShareObservations"] = new JObject { ["type"] = "array", ["items"] = share },
            });
        }

        private static MomDocument ParseDocument(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            string json = text.Trim();
            int fence = json.IndexOf("```", StringComparison.Ordinal);
            if (fence >= 0)
            {
                int start = json.IndexOf('{', fence);
                int end = json.LastIndexOf('}');
                if (start >= 0 && end > start)
                {
                    json = json.Substring(start, end - start + 1);
                }
            }
            else if (!json.StartsWith("{"))
            {
                int start = json.IndexOf('{');
                int end = json.LastIndexOf('}');
                if (start < 0 || end <= start)
                {
                    return null;
                }

                json = json.Substring(start, end - start + 1);
            }

            try
            {
                return JsonConvert.DeserializeObject<MomDocument>(json);
            }
            catch
            {
                return null;
            }
        }

        private static MomDocument Merge(MomDocument draft, MomDocument refined, string model)
        {
            if (!string.IsNullOrWhiteSpace(refined.Title)) draft.Title = refined.Title;
            if (!string.IsNullOrWhiteSpace(refined.ExecutiveSummary)) draft.ExecutiveSummary = refined.ExecutiveSummary;
            if (refined.KeyPoints?.Count > 0) draft.KeyPoints = refined.KeyPoints;
            if (refined.Topics?.Count > 0) draft.Topics = refined.Topics;
            if (refined.Decisions != null) draft.Decisions = refined.Decisions;
            if (refined.ActionItems != null) draft.ActionItems = refined.ActionItems;
            if (refined.OpenQuestions != null) draft.OpenQuestions = refined.OpenQuestions;
            if (refined.NextSteps != null) draft.NextSteps = refined.NextSteps;

            if (refined.Attendees?.Count > 0)
            {
                foreach (var a in refined.Attendees)
                {
                    var existing = draft.Attendees.FirstOrDefault(x => string.Equals(x.Name, a.Name, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                    {
                        if (!string.IsNullOrWhiteSpace(a.Role)) existing.Role = a.Role;
                    }
                    else
                    {
                        draft.Attendees.Add(a);
                    }
                }
            }

            if (refined.ScreenShareObservations?.Count > 0)
            {
                for (int i = 0; i < draft.ScreenShareObservations.Count; i++)
                {
                    var local = draft.ScreenShareObservations[i];
                    var ai = refined.ScreenShareObservations.FirstOrDefault(s =>
                                 string.Equals(s.Presenter, local.Presenter, StringComparison.OrdinalIgnoreCase) && s.StartTime == local.StartTime)
                             ?? (i < refined.ScreenShareObservations.Count ? refined.ScreenShareObservations[i] : null);
                    if (ai != null)
                    {
                        if (!string.IsNullOrWhiteSpace(ai.Summary)) local.Summary = ai.Summary;
                        if (ai.ContentObserved?.Count > 0) local.ContentObserved = ai.ContentObserved;
                    }
                }
            }

            draft.GeneratedBy = $"Claude ({model}) from transcript + {draft.ScreenShareObservations.Sum(s => s.SnapshotFiles.Count)} screen snapshots";
            draft.GeneratedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            return draft;
        }

        private static List<string> ChooseSnapshots(IList<string> all, int max)
        {
            var existing = (all ?? new List<string>()).Where(File.Exists).OrderBy(p => p).ToList();
            if (existing.Count <= max || max <= 0)
            {
                return existing;
            }

            // Evenly spaced sample over the whole meeting so every sharing period is represented.
            var chosen = new List<string>();
            double step = (double)existing.Count / max;
            for (int i = 0; i < max; i++)
            {
                chosen.Add(existing[(int)Math.Floor(i * step)]);
            }

            return chosen;
        }

        private static string LoadDownscaledJpegBase64(string path, int maxWidth, int quality)
        {
            try
            {
                using (var img = new Bitmap(path))
                {
                    if (img.Width <= maxWidth)
                    {
                        return Convert.ToBase64String(File.ReadAllBytes(path));
                    }

                    int h = (int)Math.Round(img.Height * ((double)maxWidth / img.Width));
                    using (var resized = VideoFrameConverter.ResizeBitmap(img, maxWidth, h))
                    {
                        return Convert.ToBase64String(VideoFrameConverter.EncodeJpeg(resized, quality));
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private static string Truncate(string s, int max) => string.IsNullOrEmpty(s) || s.Length <= max ? s : s.Substring(0, max) + "...";
    }
}
