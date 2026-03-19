namespace SwWatchdog;

/// <summary>
/// Result of a COM health check against the SolidWorks process.
/// Used by per-call resilience (Stage 3) to classify COMException:
/// file-level error (Healthy) vs. session degradation (Degraded/Dead).
/// </summary>
public enum SwHealthStatus
{
    /// <summary>
    /// COM channel is functional. The COMException was caused by the file/model, not SW itself.
    /// </summary>
    Healthy = 0,

    /// <summary>
    /// COM sentinel call failed. SW process is alive but COM is broken.
    /// The watchdog marks the process as degraded internally (next session boundary will restart).
    /// Caller should stop work and trigger a restart.
    /// </summary>
    Degraded = 1,

    /// <summary>
    /// SW process is not running (crashed or was killed by hang detection).
    /// Caller should trigger a restart before retrying.
    /// </summary>
    Dead = 2,
}
