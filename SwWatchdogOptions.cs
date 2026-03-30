namespace SwWatchdog;

/// <summary>
/// Configurable thresholds for watchdog behavior.
/// </summary>
public sealed class SwWatchdogOptions
{
    /// <summary>
    /// Path to sldworks.exe. If null, resolved from registry.
    /// </summary>
    public string? SolidWorksPath { get; init; }

    /// <summary>
    /// Max time to wait for SW to start (Process.Start → ROT registration → StartupProcessCompleted).
    /// </summary>
    public TimeSpan StartupTimeout { get; init; } = TimeSpan.FromSeconds(120);

    /// <summary>
    /// How often the hang detection loop runs (sends WM_NULL to SW window
    /// and checks for modal dialogs). Lower = faster detection, higher CPU.
    /// Also used as the base for exponential backoff during hang confirmation.
    /// </summary>
    public TimeSpan HangCheckInterval { get; init; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// How long SendMessageTimeoutW waits for SW to respond to WM_NULL
    /// before declaring the window unresponsive for that single check.
    /// </summary>
    public TimeSpan HangCheckTimeout { get; init; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Number of consecutive unresponsive checks before confirming hang and killing SW.
    /// Uses exponential backoff between retries (HangCheckInterval × 2^attempt).
    /// Prevents false positives during long operations (large assembly load).
    /// With default 3 retries and 5s interval: ~35 seconds total before kill.
    /// </summary>
    public int HangConfirmRetries { get; init; } = 3;

    /// <summary>
    /// Restart SW when free system memory drops below this threshold in MB (0 = disabled).
    /// Pressure formula: FreeMemoryLimitMb / actualFreeMemoryMb.
    /// </summary>
    public long FreeMemoryLimitMb { get; init; } = 512;

    /// <summary>
    /// GDI object threshold as percentage of the system GDI limit (0 = disabled).
    /// Default 80% is below SW's own 85% warning — leaves buffer for finalization.
    /// </summary>
    public int GdiObjectsLimitPercent { get; init; } = 80;

    /// <summary>
    /// USER object threshold as percentage of the system USER limit (0 = disabled).
    /// </summary>
    public int UserObjectsLimitPercent { get; init; } = 80;

    /// <summary>
    /// Kill SW process after no sessions for this duration (Zero = disabled).
    /// Frees system resources when the service is idle. SW restarts on next request.
    /// </summary>
    public TimeSpan IdleShutdownAfter { get; init; } = TimeSpan.FromMinutes(30);

    /// <summary>
    /// Polling interval for ROT (Running Object Table) during SW startup.
    /// After Process.Start, SW takes time to register its COM object in ROT.
    /// Lower = faster startup detection, higher CPU during launch.
    /// </summary>
    public TimeSpan RotPollingInterval { get; init; } = TimeSpan.FromMilliseconds(500);

    /// <summary>
    /// COM CLSID of the Bridge add-in, in braces: <c>{xxxxxxxx-xxxx-...}</c>.
    /// Used for registry startup check (HKCU\Software\SolidWorks\AddInsStartup\{GUID}).
    /// If null, add-in management is disabled — Watchdog does not verify add-in loading.
    /// </summary>
    public string? AddInClsid { get; init; }

    /// <summary>
    /// Full path to the Bridge add-in DLL.
    /// Used for <c>ISldWorks.LoadAddIn(path)</c> after SW startup to guarantee the add-in
    /// is loaded even if the startup registry key was missing or SW skipped it.
    /// Required when <see cref="AddInClsid"/> is set.
    /// </summary>
    public string? AddInDllPath { get; init; }
}
