namespace TeamsCallingBot.Http
{
    using System;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using TeamsCallingBot.Bot;

    [Route("api/calls")]
    [ApiController]
    public class CallManagementController : ControllerBase
    {
        private readonly Bot bot;

        public CallManagementController(Bot bot)
        {
            this.bot = bot ?? throw new ArgumentNullException(nameof(bot));
        }

        [HttpGet]
        public IActionResult GetActiveCalls()
        {
            var calls = this.bot.CallHandlers.Select(kvp => new
            {
                CallId = kvp.Key,
                IsMuted = kvp.Value.IsMuted,
                RecordingsFolder = kvp.Value.RecordingsManager.SessionDirectory
            });

            return this.Ok(calls);
        }

        [HttpPost("{callId}/mute")]
        public async Task<IActionResult> MuteAsync(string callId)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                await handler.MuteAsync().ConfigureAwait(false);
                return this.Ok(new { Message = "Bot muted successfully.", CallId = callId, IsMuted = true });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/unmute")]
        public async Task<IActionResult> UnmuteAsync(string callId)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                await handler.UnmuteAsync().ConfigureAwait(false);
                return this.Ok(new { Message = "Bot unmuted successfully.", CallId = callId, IsMuted = false });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/play-chime")]
        public async Task<IActionResult> PlayChimeAsync(string callId)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                await handler.AudioSender.PlayChimeAsync().ConfigureAwait(false);
                return this.Ok(new { Message = "Chime played successfully.", CallId = callId });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/play-audio")]
        public async Task<IActionResult> PlayAudioAsync(string callId, [FromBody] PlayAudioRequest request)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                if (string.IsNullOrWhiteSpace(request?.FilePath))
                {
                    return this.BadRequest(new { Message = "FilePath is required." });
                }

                await handler.PlayAudioFileAsync(request.FilePath).ConfigureAwait(false);
                return this.Ok(new { Message = "Audio playback started.", FilePath = request.FilePath, CallId = callId });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/capture-photo")]
        public IActionResult CapturePhoto(string callId, [FromQuery] string tag = "manual")
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                var photoPath = handler.CapturePhotoNow(tag);
                if (photoPath != null)
                {
                    return this.Ok(new { Message = "Photo captured successfully.", PhotoPath = photoPath, CallId = callId });
                }

                return this.BadRequest(new { Message = "No active screen share frame available to capture yet.", CallId = callId });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/send-message")]
        public async Task<IActionResult> SendChatMessageAsync(string callId, [FromBody] SendMessageRequest request)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                if (string.IsNullOrWhiteSpace(request?.Message))
                {
                    return this.BadRequest(new { Message = "Message is required." });
                }

                bool sent = await handler.PostTextMessageToChatAsync(request.Message).ConfigureAwait(false);
                if (sent)
                {
                    return this.Ok(new { Message = "Message posted to meeting chat.", CallId = callId, Content = request.Message });
                }
                else
                {
                    return this.StatusCode(502, new { Message = "Failed posting message to meeting chat (check token permissions or threadId).", CallId = callId });
                }
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/leave")]
        public async Task<IActionResult> LeaveAsync(string callId)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                await handler.Call.DeleteAsync().ConfigureAwait(false);
                return this.Ok(new { Message = "Bot left meeting.", CallId = callId });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        [HttpPost("{callId}/speak")]
        public async Task<IActionResult> SpeakAsync(string callId, [FromBody] SpeakRequest request)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                if (string.IsNullOrWhiteSpace(request?.Text))
                {
                    return this.BadRequest(new { Message = "Text is required." });
                }

                if (handler.AudioSender == null)
                {
                    return this.BadRequest(new { Message = "AudioSender is not available for this call." });
                }

                await handler.AudioSender.SpeakAsync(request.Text).ConfigureAwait(false);
                return this.Ok(new { Message = "Speech synthesized and played into meeting.", Text = request.Text, CallId = callId });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }

        // ===================================================================
        // TDA (Tata Steel Digital Assistant) integration - see CallHandler.SpeakTdaAnswerAsync /
        // SendTdaMessageAsync and Config/BotOptions.cs (TdaOptions). Both endpoints return a clear
        // 501-style message (not a 500) when TDA is not enabled/configured, since that is the
        // expected default state.
        // ===================================================================

        [HttpPost("{callId}/tda/ask")]
        public async Task<IActionResult> AskTdaAsync(string callId, [FromBody] TdaAskRequest request)
        {
            if (!this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                return this.NotFound(new { Message = $"Call {callId} not found." });
            }

            if (string.IsNullOrWhiteSpace(request?.Query))
            {
                return this.BadRequest(new { Message = "Query is required." });
            }

            bool spoken = await handler.SpeakTdaAnswerAsync(request.Query).ConfigureAwait(false);
            if (spoken)
            {
                return this.Ok(new { Message = "TDA answer spoken into meeting.", CallId = callId, Query = request.Query });
            }

            return this.StatusCode(503, new { Message = "TDA is not enabled/configured, or returned no answer. See Bot:Tda in appsettings.json.", CallId = callId });
        }

        [HttpPost("{callId}/tda/send-message")]
        public async Task<IActionResult> SendTdaMessageAsync(string callId, [FromBody] SendMessageRequest request)
        {
            if (!this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                return this.NotFound(new { Message = $"Call {callId} not found." });
            }

            if (string.IsNullOrWhiteSpace(request?.Message))
            {
                return this.BadRequest(new { Message = "Message is required." });
            }

            bool sent = await handler.SendTdaMessageAsync(request.Message).ConfigureAwait(false);
            if (sent)
            {
                return this.Ok(new { Message = "Message sent to TDA.", CallId = callId });
            }

            return this.StatusCode(503, new { Message = "TDA is not enabled/configured, or the send failed. See Bot:Tda in appsettings.json.", CallId = callId });
        }

        // ===================================================================
        // Bot video tile visualization (chart/image shown instead of the status card) - see
        // CallHandler.ShowVisualization/ClearVisualization and
        // Video/VideoFrameConverter.CreateVisualizationFrame.
        // ===================================================================

        [HttpPost("{callId}/show-visualization")]
        public IActionResult ShowVisualization(string callId, [FromBody] ShowVisualizationRequest request)
        {
            if (!this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                return this.NotFound(new { Message = $"Call {callId} not found." });
            }

            if (string.IsNullOrWhiteSpace(request?.ImageBase64))
            {
                return this.BadRequest(new { Message = "ImageBase64 (PNG/JPEG bytes, base64-encoded) is required." });
            }

            try
            {
                byte[] bytes = Convert.FromBase64String(request.ImageBase64);
                using (var ms = new MemoryStream(bytes))
                using (var image = (Bitmap)Image.FromStream(ms))
                {
                    handler.ShowVisualization(image, request.Title, request.DurationSeconds);
                }

                return this.Ok(new { Message = "Visualization is now showing on the bot's video tile.", CallId = callId, DurationSeconds = request.DurationSeconds });
            }
            catch (Exception ex)
            {
                return this.BadRequest(new { Message = $"Could not decode ImageBase64 as an image: {ex.Message}" });
            }
        }

        [HttpPost("{callId}/clear-visualization")]
        public IActionResult ClearVisualization(string callId)
        {
            if (this.bot.CallHandlers.TryGetValue(callId, out var handler))
            {
                handler.ClearVisualization();
                return this.Ok(new { Message = "Bot video tile reverted to status card.", CallId = callId });
            }

            return this.NotFound(new { Message = $"Call {callId} not found." });
        }
    }

    public class SpeakRequest
    {
        public string Text { get; set; }
    }

    public class SendMessageRequest
    {
        public string Message { get; set; }
    }

    public class PlayAudioRequest
    {
        public string FilePath { get; set; }
    }

    public class TdaAskRequest
    {
        public string Query { get; set; }
    }

    public class ShowVisualizationRequest
    {
        /// <summary>Base64-encoded PNG or JPEG image bytes.</summary>
        public string ImageBase64 { get; set; }

        public string Title { get; set; }

        /// <summary>0 or negative = show indefinitely until clear-visualization is called.</summary>
        public int DurationSeconds { get; set; } = 30;
    }
}
