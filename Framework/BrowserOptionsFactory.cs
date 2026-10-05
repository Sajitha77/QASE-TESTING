using System;
using System.Linq;
using OpenQA.Selenium.Chrome;

namespace TestAutomationFramework.Framework
{
    public static class BrowserOptionsFactory
    {
        // Keep host selection explicit. Never allow automatic sign-in to all sites.
        public static string? ValidateAuthentication(bool windowsAuth, bool incognito, string? authServerAllowlist)
        {
            if (!windowsAuth)
            {
                if (!string.IsNullOrWhiteSpace(authServerAllowlist))
                    throw new ArgumentException("--authServerAllowlist requires --windowsAuth true.");
                return null;
            }

            if (incognito)
                throw new ArgumentException("Windows authentication cannot be combined with Incognito. Use --incognito false.");

            if (string.IsNullOrWhiteSpace(authServerAllowlist))
                throw new ArgumentException("Windows authentication requires trusted hostnames in --authServerAllowlist.");

            var hosts = authServerAllowlist.Split(',').Select(host => host.Trim().TrimEnd('.').ToLowerInvariant()).ToArray();
            foreach (var host in hosts)
            {
                var kind = Uri.CheckHostName(host);
                if (string.IsNullOrEmpty(host) || host.Contains('*') || host.Any(char.IsWhiteSpace) ||
                    (kind != UriHostNameType.Dns && kind != UriHostNameType.IPv4) ||
                    host.IndexOfAny(new[] { '/', '\\', ':', '@', '?', '#', '=' }) >= 0)
                {
                    throw new ArgumentException("Enter comma-separated exact hostnames only, without https://, paths, credentials, ports or wildcards.");
                }
            }

            return string.Join(",", hosts.Distinct(StringComparer.OrdinalIgnoreCase));
        }

        public static void EnsureWindowsAuthenticationSupported(bool windowsAuth)
        {
            if (windowsAuth && !OperatingSystem.IsWindows())
                throw new PlatformNotSupportedException("Windows authentication mode must run on Windows under an account with access to the test site.");
        }

        public static ChromeOptions Create(bool headless, bool incognito, bool windowsAuth = false, string? authServerAllowlist = null)
        {
            var hosts = ValidateAuthentication(windowsAuth, incognito, authServerAllowlist);
            var options = new ChromeOptions();

            if (headless)
            {
                options.AddArgument("--headless=new");
                options.AddArgument("--window-size=1920,1080");
            }
            else
            {
                options.AddArgument("--start-maximized");
            }

            if (incognito)
                options.AddArgument("--incognito");

            // Chrome uses the Windows identity of its process. ChromeDriver creates
            // a separate temporary profile; no personal profile or credentials are read.
            if (windowsAuth)
                options.AddArgument($"--auth-server-allowlist={hosts}");

            options.AddArgument("--log-level=3");
            options.AddExcludedArgument("enable-logging");
            return options;
        }
    }
}
