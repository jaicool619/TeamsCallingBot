namespace TeamsCallingBot.Mom
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using TeamsCallingBot.Config;
    using TeamsCallingBot.Storage;

    /// <summary>
    /// Builds a complete MoM document WITHOUT any network call, purely from the artefacts the bot
    /// produced locally: the per-speaker transcript, the meeting timeline (joins/leaves, screen
    /// shares with snapshots) and the list of recording files. Heuristics (keyword patterns) pick
    /// out candidate decisions, action items and open questions. When an Anthropic API key is
    /// configured the ClaudeMomSummarizer refines this draft; otherwise this IS the final document.
    /// </summary>
    public static class LocalMomBuilder
    {
        private static readonly Regex ActionPattern = new Regex(
            @"\b(action item|action point|will (send|share|prepare|update|follow|check|review|create|schedule|circulate|complete|do|take)|needs? to|need to|have to|has to|should|must|to[- ]do|follow[- ]up|by (monday|tuesday|wednesday|thursday|friday|saturday|sunday|tomorrow|next week|end of (the )?(day|week|month)|eod|eow|eom)|assign(ed)? to|take care of|responsible for|owner)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex DecisionPattern = new Regex(
            @"\b(decided|decision|agreed|agree to|agreement|finali[sz]ed?|approved|approval|sign[- ]?off|go ahead|green light|confirmed that|we will go with|conclusion|concluded)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex OwnerPattern = new Regex(
            @"\b([A-Z][a-z]+(?: [A-Z][a-z]+)?) (will|needs to|has to|should|to|is going to)\b",
            RegexOptions.Compiled);

        private static readonly Regex DuePattern = new Regex(
            @"\b(by|before|until|on) (monday|tuesday|wednesday|thursday|friday|saturday|sunday|tomorrow|today|tonight|next week|next month|end of (the )?(day|week|month|quarter)|eod|eow|eom|\d{1,2}(st|nd|rd|th)?( of)? (jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*|(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]* \d{1,2})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static MomDocument Build(
            MeetingTimeline timeline,
            IList<TranscriptEntry> transcript,
            IDictionary<string, double> talkTimeSecondsBySpeaker,
            IList<string> recordingFiles,
            MomOptions options,
            string sessionDirectory)
        {
            options = options ?? new MomOptions();
            timeline = timeline ?? new MeetingTimeline();
            transcript = transcript ?? new List<TranscriptEntry>();

            var start = timeline.StartedAt;
            var end = timeline.EndedAt ?? DateTime.Now;

            var doc = new MomDocument
            {
                Title = $"Minutes of Meeting - {start:dd MMM yyyy}",
                Organization = options.OrganizationName,
                MeetingDate = start.ToString("dddd, dd MMMM yyyy", CultureInfo.InvariantCulture),
                StartTime = start.ToString("HH:mm"),
                EndTime = end.ToString("HH:mm"),
                DurationMinutes = Math.Round((end - start).TotalMinutes, 1),
                CallId = timeline.CallId,
                GeneratedBy = "TeamsCallingBot local builder (no AI)",
                GeneratedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            };

            // Attendees ---------------------------------------------------------------------
            foreach (var group in timeline.HumanParticipants().GroupBy(p => p.DisplayName))
            {
                var first = group.OrderBy(p => p.JoinedAt).First();
                var last = group.OrderByDescending(p => p.LeftAt ?? DateTime.MaxValue).First();
                double talk = 0;
                if (talkTimeSecondsBySpeaker != null && talkTimeSecondsBySpeaker.TryGetValue(group.Key, out var seconds))
                {
                    talk = seconds;
                }

                doc.Attendees.Add(new MomAttendee
                {
                    Name = group.Key,
                    Role = "Participant",
                    JoinedAt = first.JoinedAt.ToString("HH:mm"),
                    LeftAt = last.LeftAt.HasValue ? last.LeftAt.Value.ToString("HH:mm") : end.ToString("HH:mm"),
                    TalkTimeMinutes = Math.Round(talk / 60.0, 1),
                });
            }

            // Speakers that appear in the transcript but were never resolved to a participant
            foreach (var speaker in transcript.Select(t => t.Speaker).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct())
            {
                if (!doc.Attendees.Any(a => string.Equals(a.Name, speaker, StringComparison.OrdinalIgnoreCase)))
                {
                    double talk = 0;
                    talkTimeSecondsBySpeaker?.TryGetValue(speaker, out talk);
                    doc.Attendees.Add(new MomAttendee { Name = speaker, Role = "Speaker", TalkTimeMinutes = Math.Round(talk / 60.0, 1) });
                }
            }

            // Transcript --------------------------------------------------------------------
            foreach (var entry in transcript.OrderBy(t => t.Timestamp))
            {
                doc.Transcript.Add(new MomTranscriptLine { Timestamp = entry.Timestamp, Speaker = entry.Speaker, Text = entry.Transcript });
            }

            var sentences = SplitIntoSentences(transcript);

            // Key points / decisions / actions / questions ----------------------------------
            foreach (var s in sentences)
            {
                if (s.Text.Length < 12)
                {
                    continue;
                }

                if (DecisionPattern.IsMatch(s.Text))
                {
                    doc.Decisions.Add($"{s.Text} ({s.Speaker})");
                }

                if (ActionPattern.IsMatch(s.Text))
                {
                    var ownerMatch = OwnerPattern.Match(s.Text);
                    var dueMatch = DuePattern.Match(s.Text);
                    doc.ActionItems.Add(new MomActionItem
                    {
                        Task = s.Text,
                        Owner = ownerMatch.Success ? ownerMatch.Groups[1].Value : s.Speaker,
                        DueDate = dueMatch.Success ? dueMatch.Value : "Not specified",
                        Status = "Open",
                        Source = $"{s.Speaker} at {s.Timestamp}",
                    });
                }

                if (s.Text.TrimEnd().EndsWith("?"))
                {
                    doc.OpenQuestions.Add($"{s.Text} ({s.Speaker})");
                }
            }

            doc.Decisions = doc.Decisions.Distinct().Take(25).ToList();
            doc.ActionItems = doc.ActionItems.GroupBy(a => a.Task).Select(g => g.First()).Take(40).ToList();
            doc.OpenQuestions = doc.OpenQuestions.Distinct().Take(25).ToList();

            // Topics: split the transcript into ~10 minute blocks ---------------------------
            doc.Topics = BuildTimeBlocks(transcript, start);

            // Key points: the longest statements per speaker (proxy for substance) -----------
            doc.KeyPoints = sentences
                .Where(s => s.Text.Length > 60)
                .GroupBy(s => s.Speaker)
                .SelectMany(g => g.OrderByDescending(x => x.Text.Length).Take(2))
                .OrderBy(s => s.Timestamp)
                .Take(12)
                .Select(s => $"{s.Speaker}: {s.Text}")
                .ToList();

            // Screen share observations -----------------------------------------------------
            foreach (var share in timeline.ScreenShareSessions.OrderBy(s => s.StartedAt))
            {
                var shareEnd = share.EndedAt ?? end;
                var during = transcript
                    .Where(t => TryParse(t.Timestamp, out var ts) && ts >= share.StartedAt.AddSeconds(-5) && ts <= shareEnd.AddSeconds(5))
                    .Select(t => $"{t.Speaker}: {t.Transcript}")
                    .ToList();

                doc.ScreenShareObservations.Add(new MomScreenShare
                {
                    Presenter = share.PresenterName,
                    StartTime = share.StartedAt.ToString("HH:mm:ss"),
                    EndTime = shareEnd.ToString("HH:mm:ss"),
                    DurationMinutes = Math.Round((shareEnd - share.StartedAt).TotalMinutes, 1),
                    Summary = during.Count > 0
                        ? $"{share.PresenterName} shared their screen for {Math.Round((shareEnd - share.StartedAt).TotalMinutes, 1)} minutes. Discussion while sharing: " + string.Join(" ", during.Take(6))
                        : $"{share.PresenterName} shared their screen for {Math.Round((shareEnd - share.StartedAt).TotalMinutes, 1)} minutes. See the recording and snapshots below.",
                    VideoFiles = (share.Mp4Files != null && share.Mp4Files.Count > 0 ? share.Mp4Files : share.VideoFiles).Select(Path.GetFileName).ToList(),
                    SnapshotFiles = share.SnapshotFiles.Select(Path.GetFileName).ToList(),
                });
            }

            // Recordings --------------------------------------------------------------------
            if (recordingFiles != null)
            {
                doc.Recordings = recordingFiles.Select(Path.GetFileName).OrderBy(f => f).ToList();
            }

            // Executive summary -------------------------------------------------------------
            var summary = new StringBuilder();
            summary.Append($"The meeting took place on {doc.MeetingDate} from {doc.StartTime} to {doc.EndTime} ({doc.DurationMinutes} minutes) ");
            summary.Append($"with {doc.Attendees.Count} participant(s)");
            if (doc.Attendees.Count > 0)
            {
                summary.Append(": " + string.Join(", ", doc.Attendees.Select(a => a.Name)));
            }

            summary.Append(". ");
            if (timeline.ScreenShareSessions.Count > 0)
            {
                summary.Append($"Screen was shared {timeline.ScreenShareSessions.Count} time(s) by {string.Join(", ", timeline.ScreenShareSessions.Select(s => s.PresenterName).Distinct())}. ");
            }

            summary.Append($"{transcript.Count} transcript segment(s) were captured, yielding {doc.Decisions.Count} candidate decision(s), {doc.ActionItems.Count} candidate action item(s) and {doc.OpenQuestions.Count} open question(s). ");
            if (string.IsNullOrWhiteSpace(options.AnthropicApiKey) && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY")))
            {
                summary.Append("This document was assembled locally with keyword heuristics; configure Bot:Mom:AnthropicApiKey to enable AI-written narrative minutes that also interpret the shared screen.");
            }

            doc.ExecutiveSummary = summary.ToString();

            if (doc.NextSteps.Count == 0 && doc.ActionItems.Count > 0)
            {
                doc.NextSteps = doc.ActionItems.Take(5).Select(a => $"{a.Owner}: {Truncate(a.Task, 140)}").ToList();
            }

            return doc;
        }

        // ------------------------------------------------------------------

        private sealed class Sentence
        {
            public string Speaker;
            public string Timestamp;
            public string Text;
        }

        private static List<Sentence> SplitIntoSentences(IList<TranscriptEntry> transcript)
        {
            var result = new List<Sentence>();
            foreach (var entry in transcript)
            {
                if (string.IsNullOrWhiteSpace(entry.Transcript) || entry.Transcript.StartsWith("["))
                {
                    continue; // skip the "[Spoken audio activity detected...]" fallbacks
                }

                var parts = Regex.Split(entry.Transcript, @"(?<=[\.\!\?])\s+");
                foreach (var p in parts)
                {
                    var text = p.Trim();
                    if (text.Length > 0)
                    {
                        result.Add(new Sentence { Speaker = entry.Speaker, Timestamp = entry.Timestamp, Text = text });
                    }
                }
            }

            return result;
        }

        private static List<MomTopic> BuildTimeBlocks(IList<TranscriptEntry> transcript, DateTime meetingStart)
        {
            var topics = new List<MomTopic>();
            var usable = transcript.Where(t => !string.IsNullOrWhiteSpace(t.Transcript) && !t.Transcript.StartsWith("[")).ToList();
            if (usable.Count == 0)
            {
                return topics;
            }

            const int blockMinutes = 10;
            var groups = usable
                .Select(t => new { Entry = t, Time = TryParse(t.Timestamp, out var ts) ? ts : meetingStart })
                .GroupBy(x => (int)Math.Floor((x.Time - meetingStart).TotalMinutes / blockMinutes))
                .OrderBy(g => g.Key);

            int n = 1;
            foreach (var g in groups)
            {
                var from = meetingStart.AddMinutes(g.Key * blockMinutes);
                var to = from.AddMinutes(blockMinutes);
                var text = string.Join(" ", g.Select(x => $"{x.Entry.Speaker}: {x.Entry.Transcript}"));
                topics.Add(new MomTopic
                {
                    Title = $"Discussion block {n}",
                    TimeRange = $"{from:HH:mm} - {to:HH:mm}",
                    Discussion = Truncate(text, 2500),
                    Highlights = g.Select(x => x.Entry.Transcript).Where(t => t.Length > 40).OrderByDescending(t => t.Length).Take(3).Select(t => Truncate(t, 200)).ToList(),
                });
                n++;
            }

            return topics;
        }

        private static bool TryParse(string timestamp, out DateTime value)
        {
            return DateTime.TryParseExact(timestamp, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out value)
                   || DateTime.TryParse(timestamp, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
        }

        private static string Truncate(string text, int max)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= max)
            {
                return text;
            }

            return text.Substring(0, max - 3) + "...";
        }
    }
}
