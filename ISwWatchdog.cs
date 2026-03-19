using SolidWorks.Interop.sldworks;

namespace SwWatchdog;

/// <summary>
/// Manages a single SolidWorks instance: serialized STA access, lifecycle, hang detection, restart.
/// </summary>
public interface ISwWatchdog : IAsyncDisposable
{
    /// <summary>
    /// Execute an operation on the STA thread with access to ISldWorks.
    /// Watchdog guarantees: SW is running, STA thread, serialized access.
    /// </summary>
    Task<TResult> ExecuteAsync<TResult>(
        Func<ISldWorks, TResult> operation,
        CancellationToken ct = default
    );

    /// <summary>
    /// Fire-and-forget variant (no result).
    /// </summary>
    Task ExecuteAsync(Action<ISldWorks> operation, CancellationToken ct = default);

    /// <summary>
    /// Acquire a session: isolation protocol + SetSearchFolders(workingDir).
    /// Only one session at a time (exclusive access).
    /// </summary>
    Task<ISwSession> AcquireSessionAsync(
        string workingDirectory,
        TimeSpan timeout,
        CancellationToken ct = default
    );

    /// <summary>
    /// Current watchdog status.
    /// </summary>
    SwWatchdogStatus GetStatus();

    /// <summary>
    /// Current resource pressure across GDI, USER, and system memory.
    /// Returns the worst (highest) pressure level among all resources.
    /// Safe to call from any thread (Win32 only, no COM/STA dispatch).
    /// Returns <see cref="ResourcePressure.Low"/> if SolidWorks is not running.
    /// </summary>
    ResourcePressure GetResourcePressure();

    /// <summary>
    /// Check COM channel health using a lightweight sentinel call (RevisionNumber).
    /// <para>
    /// <b>MUST be called from the STA thread</b> (inside <c>session.ExecuteAsync</c> lambda).
    /// Does NOT dispatch to STA — caller is expected to already be on STA.
    /// </para>
    /// <para>Side effects:</para>
    /// <list type="bullet">
    ///   <item><see cref="SwHealthStatus.Degraded"/>: marks the process as degraded
    ///         (next session boundary will trigger a restart).</item>
    ///   <item><see cref="SwHealthStatus.Dead"/> / <see cref="SwHealthStatus.Healthy"/>: no side effects.</item>
    /// </list>
    /// </summary>
    SwHealthStatus CheckComHealth();
}
