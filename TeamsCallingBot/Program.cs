namespace TeamsCallingBot
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Security.Cryptography.X509Certificates;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using TeamsCallingBot.Bot;

    /// <summary>
    /// Entry point. Self-hosts via Kestrel - no Windows Service wrapper yet.
    /// Once this is proven working against a real meeting, wrap it as a Windows Service
    /// (Microsoft.AspNetCore.Hosting.WindowsServices) so it can restart cleanly on boot.
    ///
    /// NOTE (2026-09-04): HttpRouteConstants.CallRoute is currently "/api/calling/notification", but
    /// Microsoft's original Cloud Service sample registered "/callback" instead. The real Azure Bot
    /// registration's Calling Webhook URL in Azure Portal must match this - verify before chasing
    /// missing-callback bugs.
    /// </summary>
    public static class Program
    {
        public static void Main(string[] args)
        {
            var host = BuildWebHost(args);

            // AUTO-JOIN ON STARTUP (TEST HARNESS):
            // Deliberately delayed 5s so Kestrel has fully bound its listening ports first - otherwise
            // Graph's notification ping could arrive before the controller is actually listening.
            // If Bot:TestMeetingJoinUrl is left blank in appsettings.json, this task no-ops cleanly.
            //
            // KNOWN LIMITATION: If the bot is started BEFORE the win-acme cert is pasted into
            // appsettings.json, BuildWebHost will throw in LoadCertificateByThumbprint BEFORE Main even
            // gets here - and even if that were bypassed, MediaPlatformStartupScript.bat will fail inside
            // at Bot construction itself (CertificateThumbprint is still a TODO placeholder, so
            // MediaPlatform initialization has nothing real to find in the certificate store). That
            // failure is expected and NOT a bug to chase - it will resolve once appsettings.json's
            // TODO fields are filled in with the VM's real values.
            _ = Task.Run(async () =>
            {
                await Task.Delay(TimeSpan.FromSeconds(5)).ConfigureAwait(false);

                var config = host.Services.GetRequiredService<IConfiguration>();
                var testUrls = config.GetSection("Bot:TestMeetingJoinUrls").Get<List<string>>() ?? new List<string>();
                var singleUrl = config["Bot:TestMeetingJoinUrl"];
                if (!string.IsNullOrWhiteSpace(singleUrl) && !testUrls.Contains(singleUrl))
                {
                    testUrls.Insert(0, singleUrl);
                }

                var activeUrls = testUrls
                    .Where(u => !string.IsNullOrWhiteSpace(u) && !u.Contains("PLACEHOLDER") && u.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    .Take(5)
                    .ToList();

                if (activeUrls.Count == 0)
                {
                    return;
                }

                var bot = host.Services.GetRequiredService<TeamsCallingBot.Bot.Bot>();

                if (activeUrls.Count == 1)
                {
                    Console.WriteLine($">>> TEST JOIN starting for: {activeUrls[0]}");
                    try
                    {
                        var call = await bot.JoinCallAsync(activeUrls[0]).ConfigureAwait(false);
                        Console.WriteLine($">>> TEST JOIN accepted. Call id: {call.Id}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($">>> TEST JOIN FAILED: {ex.GetType().Name}: {ex.Message}");
                        Console.WriteLine(ex.ToString());
                    }
                }
                else
                {
                    Console.WriteLine($">>> [Multi-Call Test] Launching {activeUrls.Count} simultaneous meeting join(s) (Capacity: up to 5)...");
                    var joinTasks = activeUrls.Select(async (url, idx) =>
                    {
                        int meetingIndex = idx + 1;
                        try
                        {
                            Console.WriteLine($">>> [Meeting #{meetingIndex}/{activeUrls.Count}] Joining: {url}");
                            var call = await bot.JoinCallAsync(url).ConfigureAwait(false);
                            Console.WriteLine($">>> [Meeting #{meetingIndex}/{activeUrls.Count}] Successfully joined! Call ID: {call.Id}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($">>> [Meeting #{meetingIndex}/{activeUrls.Count}] Join failed: {ex.GetType().Name}: {ex.Message}");
                        }
                    });
                    await Task.WhenAll(joinTasks).ConfigureAwait(false);
                }
            });

            host.Run();
        }

        public static IWebHost BuildWebHost(string[] args)
        {
            var basePath = Directory.GetCurrentDirectory();
            if (!File.Exists(Path.Combine(basePath, "appsettings.json")))
            {
                var appBase = AppDomain.CurrentDomain.BaseDirectory;
                if (File.Exists(Path.Combine(appBase, "appsettings.json")))
                {
                    basePath = appBase;
                }
            }

            return WebHost.CreateDefaultBuilder(args)
                .UseContentRoot(basePath)
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config.SetBasePath(basePath);
                    config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    config.AddEnvironmentVariables();
                })
                .UseStartup<Startup>()
                .UseKestrel((context, options) =>
                {
                    var thumbprint = context.Configuration["Bot:CertificateThumbprint"];
                    var cert = LoadCertificateByThumbprint(thumbprint);
                    options.ConfigureHttpsDefaults(https => https.ServerCertificate = cert);
                })
                .UseUrls("https://0.0.0.0:443")
                .Build();
        }

        private static X509Certificate2 LoadCertificateByThumbprint(string thumbprint)
        {
            if (string.IsNullOrWhiteSpace(thumbprint) || thumbprint.StartsWith("TODO", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Bot:CertificateThumbprint in appsettings.json is missing/still a TODO placeholder - " +
                    "Kestrel cannot start its HTTPS listener without a real certificate. Run win-acme (see " +
                    "STEP8_WINACME.md) and paste the real thumbprint in first.");
            }

            using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
            {
                store.Open(OpenFlags.ReadOnly);
                var match = store.Certificates
                    .Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false)
                    .OfType<X509Certificate2>()
                    .FirstOrDefault();

                if (match == null)
                {
                    throw new InvalidOperationException(
                        $"No certificate with thumbprint '{thumbprint}' found in Cert:\\LocalMachine\\My. " +
                        "win-acme installs there by default - confirm the thumbprint was copied correctly.");
                }

                return match;
            }
        }
    }
}
