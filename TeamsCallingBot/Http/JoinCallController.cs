namespace TeamsCallingBot.Http
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using TeamsCallingBot.Bot;

    /// <summary>
    /// Manual trigger for testing: POST meeting join URL(s) here and the bot joins them.
    /// Supports joining single or multiple (up to 5) meetings concurrently.
    /// </summary>
    [Route("api/testjoin")]
    public class JoinCallController : Controller
    {
        private readonly TeamsCallingBot.Bot.Bot bot;

        public JoinCallController(TeamsCallingBot.Bot.Bot bot)
        {
            this.bot = bot;
        }

        [HttpPost]
        public async Task<IActionResult> JoinAsync([FromBody] JoinCallRequest request)
        {
            if (string.IsNullOrWhiteSpace(request?.MeetingJoinUrl))
            {
                return this.BadRequest(new { Message = "MeetingJoinUrl is required." });
            }

            var call = await this.bot.JoinCallAsync(request.MeetingJoinUrl).ConfigureAwait(false);
            return this.Ok(new { callId = call.Id });
        }

        /// <summary>
        /// Test endpoint to join up to 5 meetings concurrently.
        /// </summary>
        [HttpPost("batch")]
        public async Task<IActionResult> JoinBatchAsync([FromBody] JoinBatchCallsRequest request)
        {
            if (request?.MeetingJoinUrls == null || request.MeetingJoinUrls.Count == 0)
            {
                return this.BadRequest(new { Message = "MeetingJoinUrls list is required." });
            }

            var urls = request.MeetingJoinUrls.Where(u => !string.IsNullOrWhiteSpace(u)).Take(5).ToList();
            var tasks = urls.Select(async (url, idx) =>
            {
                try
                {
                    var call = await this.bot.JoinCallAsync(url).ConfigureAwait(false);
                    return new { Index = idx + 1, Url = url, Success = true, CallId = call.Id, Error = (string)null };
                }
                catch (Exception ex)
                {
                    return new { Index = idx + 1, Url = url, Success = false, CallId = (string)null, Error = ex.Message };
                }
            });

            var results = await Task.WhenAll(tasks).ConfigureAwait(false);
            return this.Ok(new { TotalRequested = urls.Count, Results = results });
        }
    }

    public class JoinCallRequest
    {
        public string MeetingJoinUrl { get; set; }
    }

    public class JoinBatchCallsRequest
    {
        public List<string> MeetingJoinUrls { get; set; } = new List<string>();
    }
}
