using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace SwWatchdog;

/// <summary>
/// DI registration for SwWatchdog.
/// </summary>
public static class SwWatchdogServiceExtensions
{
    private const string ConfigFileName = "swwatchdog.json";

    /// <summary>
    /// Add SwWatchdog services to the DI container.
    /// Loads defaults from swwatchdog.json next to the library DLL,
    /// then applies optional overrides from the caller.
    /// </summary>
    public static IServiceCollection AddSwWatchdog(
        this IServiceCollection services,
        Action<SwWatchdogOptions>? configure = null
    )
    {
        // 1. Load library's own config file (overrides code defaults)
        var assemblyDir = Path.GetDirectoryName(
            typeof(SwWatchdogServiceExtensions).Assembly.Location
        )!;
        var configPath = Path.Combine(assemblyDir, ConfigFileName);

        var config = new ConfigurationBuilder()
            .AddJsonFile(configPath, optional: true, reloadOnChange: false)
            .Build();

        services.Configure<SwWatchdogOptions>(config);

        // 2. Host overrides (highest priority)
        if (configure is not null)
            services.PostConfigure(configure);

        services.AddSingleton<ISwWatchdog, SwWatchdogImpl>();

        return services;
    }
}
