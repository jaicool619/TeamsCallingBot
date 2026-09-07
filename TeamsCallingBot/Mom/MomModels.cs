namespace TeamsCallingBot.Mom
{
    using System;
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// The structured Minutes-of-Meeting model. Both the local heuristic builder and the Claude
    /// summariser produce this shape; DocxWriter/MarkdownWriter render it. Property names are
    /// lower-case in JSON so the file can be consumed by the existing JS MoM pipeline.
    /// </summary>
    public class MomDocument
    {
        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("organization")]
        public string Organization { get; set; }

        [JsonProperty("meetingDate")]
        public string MeetingDate { get; set; }

        [JsonProperty("startTime")]
        public string StartTime { get; set; }

        [JsonProperty("endTime")]
        public string EndTime { get; set; }

        [JsonProperty("durationMinutes")]
        public double DurationMinutes { get; set; }

        [JsonProperty("callId")]
        public string CallId { get; set; }

        [JsonProperty("attendees")]
        public List<MomAttendee> Attendees { get; set; } = new List<MomAttendee>();

        [JsonProperty("executiveSummary")]
        public string ExecutiveSummary { get; set; }

        [JsonProperty("keyPoints")]
        public List<string> KeyPoints { get; set; } = new List<string>();

        [JsonProperty("topics")]
        public List<MomTopic> Topics { get; set; } = new List<MomTopic>();

        [JsonProperty("decisions")]
        public List<string> Decisions { get; set; } = new List<string>();

        [JsonProperty("actionItems")]
        public List<MomActionItem> ActionItems { get; set; } = new List<MomActionItem>();

        [JsonProperty("openQuestions")]
        public List<string> OpenQuestions { get; set; } = new List<string>();

        [JsonProperty("nextSteps")]
        public List<string> NextSteps { get; set; } = new List<string>();

        [JsonProperty("screenShareObservations")]
        public List<MomScreenShare> ScreenShareObservations { get; set; } = new List<MomScreenShare>();

        [JsonProperty("transcript")]
        public List<MomTranscriptLine> Transcript { get; set; } = new List<MomTranscriptLine>();

        [JsonProperty("recordings")]
        public List<string> Recordings { get; set; } = new List<string>();

        [JsonProperty("generatedBy")]
        public string GeneratedBy { get; set; }

        [JsonProperty("generatedAt")]
        public string GeneratedAt { get; set; }
    }

    public class MomAttendee
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("joinedAt")]
        public string JoinedAt { get; set; }

        [JsonProperty("leftAt")]
        public string LeftAt { get; set; }

        [JsonProperty("talkTimeMinutes")]
        public double TalkTimeMinutes { get; set; }
    }

    public class MomTopic
    {
        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("timeRange")]
        public string TimeRange { get; set; }

        [JsonProperty("discussion")]
        public string Discussion { get; set; }

        [JsonProperty("highlights")]
        public List<string> Highlights { get; set; } = new List<string>();
    }

    public class MomActionItem
    {
        [JsonProperty("task")]
        public string Task { get; set; }

        [JsonProperty("owner")]
        public string Owner { get; set; }

        [JsonProperty("dueDate")]
        public string DueDate { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; } = "Open";

        [JsonProperty("source")]
        public string Source { get; set; }
    }

    public class MomScreenShare
    {
        [JsonProperty("presenter")]
        public string Presenter { get; set; }

        [JsonProperty("startTime")]
        public string StartTime { get; set; }

        [JsonProperty("endTime")]
        public string EndTime { get; set; }

        [JsonProperty("durationMinutes")]
        public double DurationMinutes { get; set; }

        [JsonProperty("summary")]
        public string Summary { get; set; }

        [JsonProperty("contentObserved")]
        public List<string> ContentObserved { get; set; } = new List<string>();

        [JsonProperty("videoFiles")]
        public List<string> VideoFiles { get; set; } = new List<string>();

        [JsonProperty("snapshotFiles")]
        public List<string> SnapshotFiles { get; set; } = new List<string>();
    }

    public class MomTranscriptLine
    {
        [JsonProperty("timestamp")]
        public string Timestamp { get; set; }

        [JsonProperty("speaker")]
        public string Speaker { get; set; }

        [JsonProperty("text")]
        public string Text { get; set; }
    }
}
