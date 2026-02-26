using SolidWorks.Interop.sldworks;

namespace SwWatchdog;

/// <summary>
/// An isolated session with exclusive access to SolidWorks.
/// Dispose releases the lock and runs the isolation protocol (cleanup).
/// </summary>
public interface ISwSession : IAsyncDisposable
{
    /// <summary>
    /// Session ID for logging and diagnostics.
    /// </summary>
    string SessionId { get; }

    /// <summary>
    /// Working directory for this session.
    /// </summary>
    string WorkingDirectory { get; }

    /// <summary>
    /// Execute an operation within this session on the STA thread.
    /// </summary>
    Task<TResult> ExecuteAsync<TResult>(
        Func<ISldWorks, TResult> operation,
        CancellationToken ct = default
    );

    /// <summary>
    /// Fire-and-forget variant (no result).
    /// </summary>
    Task ExecuteAsync(Action<ISldWorks> operation, CancellationToken ct = default);
}
