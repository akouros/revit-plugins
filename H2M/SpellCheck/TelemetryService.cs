using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace H2M
{
    /// <summary>
    /// Provides fire-and-forget telemetry logging via HTTP POST to a configurable endpoint.
    /// <para>
    /// A single static <see cref="HttpClient"/> instance is shared across all calls
    /// (per Microsoft best practice).  All exceptions are silently swallowed so
    /// telemetry failures never surface to the user.
    /// </para>
    /// <para>
    /// <b>Usage from any command:</b>
    /// <code>
    /// var telemetry = new TelemetryService();
    /// telemetry.TrackEvent("SpellCheck", "button_click");
    /// </code>
    /// </para>
    /// </summary>
    public class TelemetryService
    {
        // Single shared HttpClient — never disposed; lives for the process lifetime.
        private static readonly HttpClient _httpClient =
            new HttpClient { Timeout = TimeSpan.FromSeconds(10) };

        /// <summary>
        /// The app.config key used to look up the telemetry endpoint URL.
        /// Set this key in <c>app.config</c> under <c>&lt;appSettings&gt;</c>.
        /// If absent or empty, telemetry calls are silently skipped.
        /// </summary>
        public const string EndpointConfigKey = "TelemetryEndpoint";

        private static readonly string _endpoint;

        static TelemetryService()
        {
            try
            {
                _endpoint = System.Configuration.ConfigurationManager
                    .AppSettings[EndpointConfigKey];
            }
            catch
            {
                _endpoint = null;
            }
        }

        /// <summary>
        /// Fires a telemetry event asynchronously without blocking the calling thread.
        /// The call is fire-and-forget: the result is never awaited and all exceptions
        /// are silently suppressed.
        /// </summary>
        /// <param name="tool">
        /// The tool name included in the payload, e.g. <c>"SpellCheck"</c>.
        /// </param>
        /// <param name="eventType">
        /// One of <c>"button_click"</c> or <c>"task_completed"</c>.
        /// </param>
        /// <param name="metadata">
        /// Optional dictionary of additional key-value pairs to include in the
        /// <c>metadata</c> field of the payload.  Pass <c>null</c> to omit.
        /// </param>
        public void TrackEvent(string tool, string eventType,
            Dictionary<string, object> metadata = null)
        {
            if (string.IsNullOrWhiteSpace(_endpoint)) return;

            // Capture values before entering the background task lambda.
            string user      = Environment.UserName;
            string machine   = Environment.MachineName;
            string timestamp = DateTime.UtcNow.ToString("o");   // ISO 8601

            var payload = new
            {
                tool        = tool,
                event_type  = eventType,
                user        = user,
                machine     = machine,
                timestamp   = timestamp,
                metadata    = metadata ?? new Dictionary<string, object>()
            };

            string json;
            try
            {
                json = JsonSerializer.Serialize(payload);
            }
            catch
            {
                return;
            }

            // Fire and forget — never await, never propagate.
            Task.Run(async () =>
            {
                try
                {
                    using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                    {
                        await _httpClient.PostAsync(_endpoint, content).ConfigureAwait(false);
                    }
                }
                catch
                {
                    // Intentionally swallowed — telemetry must never affect the host.
                }
            });
        }
    }
}
