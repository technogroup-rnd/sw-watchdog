using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SolidWorks.Interop.sldworks;
using SwWatchdog.Process;
using SwWatchdog.Sta;

namespace SwWatchdog.Session;

/// <summary>
/// Manages exclusive session access via SemaphoreSlim(1,1) with timeout.
/// </summary>
internal sealed class SessionManager : IDisposable
{
    private readonly SemaphoreSlim _semaphore = new(1, 1);
    private readonly SwWatchdogOptions _options;
    private readonly StaThread _staThread;
    private readonly IsolationProtocol _isolation;
    private readonly SwProcessManager _processManager;
    private readonly ILogger _logger;
    private readonly Lock _lock = new();

    private string? _activeSessionId;
    private bool _disposed;

    public SessionManager(
        IOptions<SwWatchdogOptions> options,
        StaThread staThread,
        IsolationProtocol isolation,
        SwProcessManager processManager,
        ILogger logger
    )
    {
        _options = options.Value;
        _staThread = staThread;
        _isolation = isolation;
        _processManager = processManager;
        _logger = logger;
    }

    public string? ActiveSessionId
    {
        get
        {
            lock (_lock)
            {
                return _activeSessionId;
            }
        }
    }

    /// <summary>
    /// Acquire exclusive session. Blocks until lock obtained or timeout.
    /// </summary>
    public async Task<SwSession> AcquireAsync(
        string workingDirectory,
        TimeSpan timeout,
        CancellationToken ct
    )
    {
        if (!await _semaphore.WaitAsync(timeout, ct))
        {
            throw new TimeoutException(
                $"Could not acquire session within {timeout}. Active session: {_activeSessionId}"
            );
        }

        var sessionId = Guid.NewGuid().ToString("N")[..12];

        try
        {
            // Ensure SW is running (launches if needed) — synchronous on STA
            await _staThread.EnqueueAsync(
                () => _processManager.EnsureRunning(CancellationToken.None),
                ct
            );

            // Check if periodic restart is needed (between sessions)
            if (_processManager.NeedsRestart())
            {
                await _staThread.EnqueueAsync(
                    () => _processManager.Restart(CancellationToken.None),
                    ct
                );
            }

            // Run isolation protocol on STA thread
            var clean = await _staThread.EnqueueAsync(
                () =>
                {
                    var swApp =
                        _processManager.SwApp
                        ?? throw new InvalidOperationException(
                            "SolidWorks not available after launch"
                        );
                    return _isolation.Isolate(swApp, workingDirectory);
                },
                ct
            );

            if (!clean)
            {
                _processManager.MarkDegraded();
                // Restart and retry isolation
                await _staThread.EnqueueAsync(
                    () => _processManager.Restart(CancellationToken.None),
                    ct
                );

                clean = await _staThread.EnqueueAsync(
                    () =>
                    {
                        var swApp =
                            _processManager.SwApp
                            ?? throw new InvalidOperationException(
                                "SolidWorks not available after restart"
                            );
                        return _isolation.Isolate(swApp, workingDirectory);
                    },
                    ct
                );

                if (!clean)
                    throw new InvalidOperationException(
                        "SolidWorks is degraded even after restart"
                    );
            }

            lock (_lock)
            {
                _activeSessionId = sessionId;
            }

            _logger.LogInformation(
                "Session {SessionId} acquired, workingDir={Dir}",
                sessionId,
                workingDirectory
            );

            return new SwSession(sessionId, workingDirectory, _staThread, _processManager, this);
        }
        catch
        {
            _semaphore.Release();
            throw;
        }
    }

    /// <summary>
    /// Release the session lock and run cleanup isolation.
    /// </summary>
    internal async Task ReleaseAsync(string sessionId)
    {
        try
        {
            // Run cleanup isolation on STA thread
            if (_processManager.IsRunning)
            {
                var clean = await _staThread.EnqueueAsync(
                    () =>
                    {
                        var swApp = _processManager.SwApp;
                        return swApp is not null && _isolation.Cleanup(swApp);
                    },
                    CancellationToken.None
                );

                if (!clean)
                    _processManager.MarkDegraded();
            }

            _logger.LogInformation("Session {SessionId} released", sessionId);
        }
        finally
        {
            lock (_lock)
            {
                _activeSessionId = null;
            }
            _semaphore.Release();
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;

        _semaphore.Dispose();
    }
}

/// <summary>
/// Active session — holds exclusive access to SolidWorks.
/// Dispose releases the lock and runs cleanup.
/// </summary>
internal sealed class SwSession : ISwSession
{
    private readonly StaThread _staThread;
    private readonly SwProcessManager _processManager;
    private readonly SessionManager _sessionManager;
    private int _disposed;

    internal SwSession(
        string sessionId,
        string workingDirectory,
        StaThread staThread,
        SwProcessManager processManager,
        SessionManager sessionManager
    )
    {
        SessionId = sessionId;
        WorkingDirectory = workingDirectory;
        _staThread = staThread;
        _processManager = processManager;
        _sessionManager = sessionManager;
    }

    public string SessionId { get; }
    public string WorkingDirectory { get; }

    public async Task<TResult> ExecuteAsync<TResult>(
        Func<ISldWorks, TResult> operation,
        CancellationToken ct
    )
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, this);

        try
        {
            return await _staThread.EnqueueAsync(
                () =>
                {
                    var swApp =
                        _processManager.SwApp
                        ?? throw new InvalidOperationException("SolidWorks is not running");
                    return operation(swApp);
                },
                ct
            );
        }
        catch (Exception ex)
            when (IsProcessKilledException(
                    ex,
                    _processManager.IsDegraded,
                    _processManager.IsRunning
                )
            )
        {
            throw new SwProcessKilledException(
                "SolidWorks was killed during active session — transient, retry with a new session",
                ex
            );
        }
    }

    public async Task ExecuteAsync(Action<ISldWorks> operation, CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, this);

        try
        {
            await _staThread.EnqueueAsync(
                () =>
                {
                    var swApp =
                        _processManager.SwApp
                        ?? throw new InvalidOperationException("SolidWorks is not running");
                    operation(swApp);
                },
                ct
            );
        }
        catch (Exception ex)
            when (IsProcessKilledException(
                    ex,
                    _processManager.IsDegraded,
                    _processManager.IsRunning
                )
            )
        {
            throw new SwProcessKilledException(
                "SolidWorks was killed during active session — transient, retry with a new session",
                ex
            );
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;

        await _sessionManager.ReleaseAsync(SessionId);
    }

    /// <summary>
    /// Determines if a COM exception should be wrapped as <see cref="SwProcessKilledException"/>.
    /// True when: exception is COMException/SEHException AND process was killed by hang detection
    /// (degraded + not running). Testable static method — called from exception filter.
    /// </summary>
    internal static bool IsProcessKilledException(Exception ex, bool isDegraded, bool isRunning) =>
        ex is COMException or SEHException && isDegraded && !isRunning;
}
