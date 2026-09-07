namespace TeamsCallingBot.Storage
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using Newtonsoft.Json;

    /// <summary>
    /// In-memory record of everything that happened in a meeting that is NOT audio: who joined and
    /// left, who shared their screen (and where the recording + snapshots ended up), which camera
    /// streams were recorded, and bot-side events (greeting spoken, muted, chat posted...).
    /// Serialised to 09_meeting_timeline.json at call end and consumed by the MoM generator.
    /// All members are thread-safe; media callbacks and Graph notifications arrive on different threads.
    /// </summary>
    public class MeetingTimeline
    {
        private readonly object sync = new object();

        public string CallId { get; set; }

        public string ChatThreadId { get; set; }

        public DateTime StartedAt { get; set; } = DateTime.Now;

        public DateTime? EndedAt { get; set; }

        public List<ParticipantRecord> Participants { get; } = new List<ParticipantRecord>();

        public List<MediaSessionRecord> ScreenShareSessions { get; } = new List<MediaSessionRecord>();

        public List<MediaSessionRecord> CameraSessions { get; } = new List<MediaSessionRecord>();

        public List<TimelineEvent> Events { get; } = new List<TimelineEvent>();

        public void AddEvent(string type, string detail)
        {
            lock (this.sync)
            {
                this.Events.Add(new TimelineEvent { At = DateTime.Now, Type = type, Detail = detail });
            }
        }

        public ParticipantRecord ParticipantJoined(string participantId, string displayName, string userId, bool isBot)
        {
            lock (this.sync)
            {
                var existing = this.Participants.FirstOrDefault(p => p.ParticipantId == participantId && p.LeftAt == null);
                if (existing != null)
                {
                    if (!string.IsNullOrWhiteSpace(displayName))
                    {
                        existing.DisplayName = displayName;
                    }

                    return existing;
                }

                var record = new ParticipantRecord
                {
                    ParticipantId = participantId,
                    DisplayName = string.IsNullOrWhiteSpace(displayName) ? "Unknown participant" : displayName,
                    UserId = userId,
                    IsBot = isBot,
                    JoinedAt = DateTime.Now,
                };
                this.Participants.Add(record);
                this.Events.Add(new TimelineEvent { At = record.JoinedAt, Type = "participant_joined", Detail = record.DisplayName });
                return record;
            }
        }

        public void ParticipantLeft(string participantId)
        {
            lock (this.sync)
            {
                var existing = this.Participants.FirstOrDefault(p => p.ParticipantId == participantId && p.LeftAt == null);
                if (existing != null)
                {
                    existing.LeftAt = DateTime.Now;
                    this.Events.Add(new TimelineEvent { At = existing.LeftAt.Value, Type = "participant_left", Detail = existing.DisplayName });
                }
            }
        }

        public MediaSessionRecord StartScreenShare(uint msi, string presenterName, string participantId)
        {
            lock (this.sync)
            {
                var record = new MediaSessionRecord
                {
                    MediaSourceId = msi,
                    PresenterName = presenterName,
                    ParticipantId = participantId,
                    StartedAt = DateTime.Now,
                };
                this.ScreenShareSessions.Add(record);
                this.Events.Add(new TimelineEvent { At = record.StartedAt, Type = "screen_share_started", Detail = presenterName });
                return record;
            }
        }

        public MediaSessionRecord StartCameraSession(uint msi, string participantName, string participantId)
        {
            lock (this.sync)
            {
                var record = new MediaSessionRecord
                {
                    MediaSourceId = msi,
                    PresenterName = participantName,
                    ParticipantId = participantId,
                    StartedAt = DateTime.Now,
                };
                this.CameraSessions.Add(record);
                this.Events.Add(new TimelineEvent { At = record.StartedAt, Type = "camera_recording_started", Detail = participantName });
                return record;
            }
        }

        public void EndMediaSession(MediaSessionRecord record, IEnumerable<string> videoFiles, IEnumerable<string> snapshotFiles, long framesWritten)
        {
            if (record == null)
            {
                return;
            }

            lock (this.sync)
            {
                record.EndedAt = DateTime.Now;
                record.FramesWritten = framesWritten;
                if (videoFiles != null)
                {
                    record.VideoFiles = videoFiles.ToList();
                }

                if (snapshotFiles != null)
                {
                    record.SnapshotFiles = snapshotFiles.ToList();
                }

                this.Events.Add(new TimelineEvent
                {
                    At = record.EndedAt.Value,
                    Type = this.ScreenShareSessions.Contains(record) ? "screen_share_ended" : "camera_recording_ended",
                    Detail = record.PresenterName,
                });
            }
        }

        public List<ParticipantRecord> HumanParticipants()
        {
            lock (this.sync)
            {
                return this.Participants.Where(p => !p.IsBot).ToList();
            }
        }

        public string Save(string directory)
        {
            var path = Path.Combine(directory, "09_meeting_timeline.json");
            lock (this.sync)
            {
                File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented), Encoding.UTF8);
            }

            return path;
        }

        public static MeetingTimeline Load(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            var loaded = JsonConvert.DeserializeObject<MeetingTimeline>(File.ReadAllText(path));
            return loaded;
        }
    }

    public class ParticipantRecord
    {
        public string ParticipantId { get; set; }

        public string DisplayName { get; set; }

        public string UserId { get; set; }

        public bool IsBot { get; set; }

        public DateTime JoinedAt { get; set; }

        public DateTime? LeftAt { get; set; }

        /// <summary>Seconds of speech attributed to this participant (filled from the per-speaker WAVs).</summary>
        public double TalkTimeSeconds { get; set; }
    }

    public class MediaSessionRecord
    {
        public uint MediaSourceId { get; set; }

        public string PresenterName { get; set; }

        public string ParticipantId { get; set; }

        public DateTime StartedAt { get; set; }

        public DateTime? EndedAt { get; set; }

        public long FramesWritten { get; set; }

        public List<string> VideoFiles { get; set; } = new List<string>();

        public List<string> Mp4Files { get; set; } = new List<string>();

        public List<string> SnapshotFiles { get; set; } = new List<string>();

        [JsonIgnore]
        public TimeSpan Duration => (this.EndedAt ?? DateTime.Now) - this.StartedAt;
    }

    public class TimelineEvent
    {
        public DateTime At { get; set; }

        public string Type { get; set; }

        public string Detail { get; set; }
    }
}
