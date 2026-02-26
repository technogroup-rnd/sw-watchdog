namespace SwWatchdog;

/// <summary>
/// Configurable thresholds for watchdog behavior.
/// </summary>
public sealed class SwWatchdogOptions
{
    /// <summary>
    /// Path to sldworks.exe. If null, resolved from registry.
    /// </summary>
    public string? SolidWorksPath { get; set; }

    /// <summary>
    /// Max time to wait for SW to start and appear in ROT.
    /// </summary>
    public TimeSpan StartupTimeout { get; set; } = TimeSpan.FromSeconds(120);

    /// <summary>
    /// Interval for hang detection polling.
    /// </summary>
    public TimeSpan HangCheckInterval { get; set; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Timeout for SendMessageTimeout hang check.
    /// </summary>
    public TimeSpan HangCheckTimeout { get; set; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Restart SW after processing this many files (0 = disabled).
    /// </summary>
    public int RestartAfterFileCount { get; set; } = 500;

    /// <summary>
    /// Restart SW after this elapsed time (Zero = disabled).
    /// </summary>
    public TimeSpan RestartAfterElapsed { get; set; } = TimeSpan.FromHours(4);

    /// <summary>
    /// Restart SW when memory exceeds this threshold in MB (0 = disabled).
    /// </summary>
    public long RestartAfterMemoryMb { get; set; } = 4096;

    /// <summary>
    /// Shutdown SW after idle for this duration (Zero = disabled).
    /// </summary>
    public TimeSpan IdleShutdownAfter { get; set; } = TimeSpan.FromMinutes(30);

    /// <summary>
    /// Max time a session can be held before forced release.
    /// </summary>
    public TimeSpan SessionWatchdogTimeout { get; set; } = TimeSpan.FromMinutes(30);

    /// <summary>
    /// ROT polling interval during SW startup.
    /// </summary>
    public TimeSpan RotPollingInterval { get; set; } = TimeSpan.FromMilliseconds(500);
}
