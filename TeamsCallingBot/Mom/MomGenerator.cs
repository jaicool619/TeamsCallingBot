namespace TeamsCallingBot.Mom
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Newtonsoft.Json;
    using TeamsCallingBot.Config;
    using TeamsCallingBot.Storage;

    /// <summary>
    /// Orchestrates MoM creation at call end (or on demand from a saved session folder):
    ///   1. LocalMomBuilder  -> structured draft from transcript + timeline (always, offline)
    ///   2. ClaudeMomSummarizer -> detailed narrative using transcript + screen snapshots (optional)
    ///   3. Writers -> 08_minutes_of_meeting.docx / .md / .json in the session folder
    /// </summary>
    public static class MomGenerator
    {
        public const string DocxFileName = "08_minutes_of_meeting.docx";
        public const string MarkdownFileName = "08_minutes_of_meeting.md";
        public const string JsonFileName = "08_minutes_of_meeting.json";

        public sealed class Result
        {
            public MomDocument Document { get; set; }

            public string DocxPath { get; set; }

            public string MarkdownPath { get; set; }

            public string JsonPath { get; set; }

            public bool UsedAi { get; set; }
        }

        public static async Task<Result> GenerateAsync(
            string sessionDirectory,
            MeetingTimeline timeline,
            IList<TranscriptEntry> transcript,
            IDictionary<string, double> talkTimeSecondsBySpeaker,
            MomOptions options,
            Action<string> log)
        {
            options = options ?? new MomOptions();
            log = log ?? (_ => { });

            var recordingFiles = Directory.Exists(sessionDirectory)
                ? Directory.GetFiles(sessionDirectory).Where(f => f.EndsWith(".wav", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".avi", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase)).ToList()
                : new List<string>();

            var doc = LocalMomBuilder.Build(timeline, transcript, talkTimeSecondsBySpeaker, recordingFiles, options, sessionDirectory);
            log($"[MoM] Local draft built: {doc.Attendees.Count} attendees, {doc.Transcript.Count} transcript lines, {doc.ScreenShareObservations.Count} screen-share sessions.");

            var result = new Result { Document = doc };

            var summarizer = new ClaudeMomSummarizer(options, log);
            if (summarizer.IsConfigured)
            {
                var snapshots = (timeline?.ScreenShareSessions ?? new List<MediaSessionRecord>())
                    .SelectMany(s => s.SnapshotFiles ?? new List<string>())
                    .Where(File.Exists)
                    .ToList();

                if (snapshots.Count == 0 && Directory.Exists(sessionDirectory))
                {
                    snapshots = Directory.GetFiles(sessionDirectory, "03_snapshot_screenshare_*.jpg").OrderBy(f => f).ToList();
                }

                var before = doc.GeneratedBy;
                doc = await summarizer.SummarizeAsync(doc, snapshots).ConfigureAwait(false);
                result.Document = doc;
                result.UsedAi = doc.GeneratedBy != before;
                log(result.UsedAi ? "[MoM] AI summarisation applied." : "[MoM] AI summarisation not applied (see earlier log lines).");
            }
            else
            {
                log("[MoM] No Anthropic API key configured - writing the local heuristic MoM only.");
            }

            Directory.CreateDirectory(sessionDirectory);
            result.JsonPath = Path.Combine(sessionDirectory, JsonFileName);
            File.WriteAllText(result.JsonPath, JsonConvert.SerializeObject(doc, Formatting.Indented), new UTF8Encoding(false));

            result.MarkdownPath = Path.Combine(sessionDirectory, MarkdownFileName);
            File.WriteAllText(result.MarkdownPath, RenderMarkdown(doc), new UTF8Encoding(false));

            if (options.GenerateWordDocument)
            {
                try
                {
                    result.DocxPath = Path.Combine(sessionDirectory, DocxFileName);
                    RenderDocx(doc, timeline, options, sessionDirectory, result.DocxPath);
                    log($"[MoM] Word document written: {result.DocxPath}");
                }
                catch (Exception ex)
                {
                    log($"[MoM] Failed writing Word document: {ex.Message}");
                    result.DocxPath = null;
                }
            }

            return result;
        }

        /// <summary>
        /// Re-generates the MoM from a finished session folder (04_transcript.json + 09_meeting_timeline.json).
        /// Used by the HTTP API so an operator can add an API key later and regenerate a richer document.
        /// </summary>
        public static Task<Result> GenerateFromSessionFolderAsync(string sessionDirectory, MomOptions options, Action<string> log)
        {
            var transcriptPath = Path.Combine(sessionDirectory, "04_transcript.json");
            var timelinePath = Path.Combine(sessionDirectory, "09_meeting_timeline.json");

            var transcript = File.Exists(transcriptPath)
                ? JsonConvert.DeserializeObject<List<TranscriptEntry>>(File.ReadAllText(transcriptPath)) ?? new List<TranscriptEntry>()
                : new List<TranscriptEntry>();

            var timeline = MeetingTimeline.Load(timelinePath) ?? new MeetingTimeline { CallId = Path.GetFileName(sessionDirectory) };

            var talkTime = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var wav in Directory.GetFiles(sessionDirectory, "02_audio_speaker_*.wav"))
            {
                // 16 kHz, 16-bit mono = 32,000 bytes/second
                double seconds = Math.Max(0, new FileInfo(wav).Length - 44) / 32000.0;
                var speaker = Path.GetFileNameWithoutExtension(wav).Replace("02_audio_", string.Empty);
                talkTime[speaker] = seconds;
            }

            return GenerateAsync(sessionDirectory, timeline, transcript, talkTime, options, log);
        }

        // ------------------------------------------------------------------

        public static string RenderMarkdown(MomDocument doc)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"# {doc.Title}");
            if (!string.IsNullOrWhiteSpace(doc.Organization))
            {
                sb.AppendLine($"**Organisation:** {doc.Organization}  ");
            }

            sb.AppendLine($"**Date:** {doc.MeetingDate}  ");
            sb.AppendLine($"**Time:** {doc.StartTime} - {doc.EndTime} ({doc.DurationMinutes} min)  ");
            sb.AppendLine($"**Call ID:** {doc.CallId}  ");
            sb.AppendLine($"**Generated:** {doc.GeneratedAt} by {doc.GeneratedBy}");
            sb.AppendLine();

            sb.AppendLine("## Attendees");
            sb.AppendLine("| Name | Role | Joined | Left | Talk time (min) |");
            sb.AppendLine("|---|---|---|---|---|");
            foreach (var a in doc.Attendees)
            {
                sb.AppendLine($"| {a.Name} | {a.Role} | {a.JoinedAt} | {a.LeftAt} | {a.TalkTimeMinutes} |");
            }

            sb.AppendLine();
            sb.AppendLine("## Executive summary");
            sb.AppendLine(doc.ExecutiveSummary);
            sb.AppendLine();

            if (doc.KeyPoints.Count > 0)
            {
                sb.AppendLine("## Key points");
                foreach (var k in doc.KeyPoints) sb.AppendLine($"- {k}");
                sb.AppendLine();
            }

            if (doc.Topics.Count > 0)
            {
                sb.AppendLine("## Discussion by topic");
                foreach (var t in doc.Topics)
                {
                    sb.AppendLine($"### {t.Title} ({t.TimeRange})");
                    sb.AppendLine(t.Discussion);
                    foreach (var h in t.Highlights) sb.AppendLine($"- {h}");
                    sb.AppendLine();
                }
            }

            sb.AppendLine("## Decisions");
            if (doc.Decisions.Count == 0) sb.AppendLine("_None recorded._");
            foreach (var d in doc.Decisions) sb.AppendLine($"- {d}");
            sb.AppendLine();

            sb.AppendLine("## Action items");
            if (doc.ActionItems.Count == 0)
            {
                sb.AppendLine("_None recorded._");
            }
            else
            {
                sb.AppendLine("| # | Task | Owner | Due | Status |");
                sb.AppendLine("|---|---|---|---|---|");
                int i = 1;
                foreach (var a in doc.ActionItems) sb.AppendLine($"| {i++} | {a.Task} | {a.Owner} | {a.DueDate} | {a.Status} |");
            }

            sb.AppendLine();

            if (doc.OpenQuestions.Count > 0)
            {
                sb.AppendLine("## Open questions");
                foreach (var q in doc.OpenQuestions) sb.AppendLine($"- {q}");
                sb.AppendLine();
            }

            if (doc.NextSteps.Count > 0)
            {
                sb.AppendLine("## Next steps");
                foreach (var n in doc.NextSteps) sb.AppendLine($"- {n}");
                sb.AppendLine();
            }

            sb.AppendLine("## Screen sharing");
            if (doc.ScreenShareObservations.Count == 0) sb.AppendLine("_No screen was shared._");
            foreach (var s in doc.ScreenShareObservations)
            {
                sb.AppendLine($"### {s.Presenter} ({s.StartTime} - {s.EndTime}, {s.DurationMinutes} min)");
                sb.AppendLine(s.Summary);
                foreach (var c in s.ContentObserved) sb.AppendLine($"- {c}");
                if (s.VideoFiles.Count > 0) sb.AppendLine($"Video: {string.Join(", ", s.VideoFiles)}");
                if (s.SnapshotFiles.Count > 0) sb.AppendLine($"Snapshots: {string.Join(", ", s.SnapshotFiles)}");
                sb.AppendLine();
            }

            sb.AppendLine("## Recordings");
            foreach (var r in doc.Recordings) sb.AppendLine($"- {r}");
            sb.AppendLine();

            sb.AppendLine("## Full transcript");
            foreach (var line in doc.Transcript) sb.AppendLine($"- `{line.Timestamp}` **{line.Speaker}:** {line.Text}");

            return sb.ToString();
        }

        public static void RenderDocx(MomDocument doc, MeetingTimeline timeline, MomOptions options, string sessionDirectory, string outputPath)
        {
            var w = new DocxWriter();
            w.AddTitle(doc.Title ?? "Minutes of Meeting");
            if (!string.IsNullOrWhiteSpace(doc.Organization))
            {
                w.AddParagraph(doc.Organization, bold: true, color: "374151");
            }

            w.AddTable(
                new[] { "Date", "Time", "Duration", "Call ID" },
                new List<IList<string>> { new[] { doc.MeetingDate, $"{doc.StartTime} - {doc.EndTime}", $"{doc.DurationMinutes} min", doc.CallId ?? string.Empty } });

            w.AddHeading("1. Attendees");
            w.AddTable(
                new[] { "Name", "Role", "Joined", "Left", "Talk time (min)" },
                doc.Attendees.Select(a => (IList<string>)new[] { a.Name, a.Role, a.JoinedAt, a.LeftAt, a.TalkTimeMinutes.ToString("0.0") }).ToList());

            w.AddHeading("2. Executive summary");
            w.AddParagraph(doc.ExecutiveSummary ?? string.Empty);

            if (doc.KeyPoints.Count > 0)
            {
                w.AddHeading("3. Key points");
                w.AddBullets(doc.KeyPoints);
            }

            if (doc.Topics.Count > 0)
            {
                w.AddHeading("4. Discussion by topic");
                foreach (var t in doc.Topics)
                {
                    w.AddHeading($"{t.Title}  ({t.TimeRange})", 2);
                    w.AddParagraph(t.Discussion ?? string.Empty);
                    if (t.Highlights.Count > 0)
                    {
                        w.AddBullets(t.Highlights);
                    }
                }
            }

            w.AddHeading("5. Decisions");
            if (doc.Decisions.Count == 0) w.AddParagraph("None recorded.", italic: true); else w.AddNumbered(doc.Decisions);

            w.AddHeading("6. Action items");
            if (doc.ActionItems.Count == 0)
            {
                w.AddParagraph("None recorded.", italic: true);
            }
            else
            {
                int i = 1;
                w.AddTable(
                    new[] { "#", "Task", "Owner", "Due", "Status" },
                    doc.ActionItems.Select(a => (IList<string>)new[] { (i++).ToString(), a.Task, a.Owner, a.DueDate, a.Status }).ToList());
            }

            if (doc.OpenQuestions.Count > 0)
            {
                w.AddHeading("7. Open questions");
                w.AddBullets(doc.OpenQuestions);
            }

            if (doc.NextSteps.Count > 0)
            {
                w.AddHeading("8. Next steps");
                w.AddBullets(doc.NextSteps);
            }

            w.AddHeading("9. Screen sharing");
            if (doc.ScreenShareObservations.Count == 0)
            {
                w.AddParagraph("No screen was shared during this meeting.", italic: true);
            }

            int embedded = 0;
            foreach (var s in doc.ScreenShareObservations)
            {
                w.AddHeading($"{s.Presenter}  ({s.StartTime} - {s.EndTime}, {s.DurationMinutes} min)", 2);
                w.AddParagraph(s.Summary ?? string.Empty);
                if (s.ContentObserved.Count > 0)
                {
                    w.AddBullets(s.ContentObserved);
                }

                if (s.VideoFiles.Count > 0)
                {
                    w.AddLabelValue("Recording", string.Join(", ", s.VideoFiles));
                }

                if (options.EmbedSnapshots)
                {
                    var perSession = Math.Max(1, options.MaxSnapshotsInDocument / Math.Max(1, doc.ScreenShareObservations.Count));
                    foreach (var snap in SpreadSample(s.SnapshotFiles, perSession))
                    {
                        if (embedded >= options.MaxSnapshotsInDocument)
                        {
                            break;
                        }

                        string full = Path.IsPathRooted(snap) ? snap : Path.Combine(sessionDirectory, snap);
                        if (w.AddImage(full, 6.2, $"{s.Presenter} - {Path.GetFileName(snap)}"))
                        {
                            embedded++;
                        }
                    }
                }
            }

            w.AddHeading("10. Recordings and artefacts");
            w.AddBullets(doc.Recordings);

            w.AddPageBreak();
            w.AddHeading("Appendix A - Full transcript");
            foreach (var line in doc.Transcript)
            {
                w.AddParagraph($"[{line.Timestamp}] {line.Speaker}: {line.Text}", sizeHalfPoints: 19);
            }

            w.AddParagraph(string.Empty);
            w.AddParagraph($"Generated {doc.GeneratedAt} by {doc.GeneratedBy}.", italic: true, color: "6B7280", sizeHalfPoints: 18);
            w.Save(outputPath);
        }

        private static IEnumerable<string> SpreadSample(IList<string> items, int max)
        {
            if (items == null || items.Count == 0)
            {
                yield break;
            }

            if (items.Count <= max)
            {
                foreach (var i in items) yield return i;
                yield break;
            }

            double step = (double)items.Count / max;
            for (int i = 0; i < max; i++)
            {
                yield return items[(int)Math.Floor(i * step)];
            }
        }
    }
}
