using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SolidWorks.Interop.sldworks;
using SwWatchdog.Process;
using SwWatchdog.Session;
using SwWatchdog.Sta;

namespace SwWatchdog;

/// <summary>
/// Main ISwWatchdog implementation. Wires up STA thread, process manager,
/// isolation protocol, session manager, and hang detection loop.
/// </summary>
internal sealed class SwWatchdogImpl : ISwWatchdog
{
    private readonly StaThread _staThread;
    private readonly SwProcessManager _processManager;
    private readonly SessionManager _sessionManager;
    private readonly SwWatchdogOptions _options;
    private readonly ILogger<SwWatchdogImpl> _logger;

    private readonly CancellationTokenSource _cts = new();
    private readonly Task _hangDetectionTask;
    private readonly Task _idleShutdownTask;
    private int _disposed;

    public SwWatchdogImpl(IOptions<SwWatchdogOptions> options, ILoggerFactory loggerFactory)
    {
        _options = options.Value;
        _logger = loggerFactory.CreateLogger<SwWatchdogImpl>();

        var staLogger = loggerFactory.CreateLogger<StaThread>();
        var processLogger = loggerFactory.CreateLogger<SwProcessManager>();
        var isolationLogger = loggerFactory.CreateLogger<IsolationProtocol>();
        var sessionLogger = loggerFactory.CreateLogger<SessionManager>();

        _staThread = new StaThread(staLogger);
        _processManager = new SwProcessManager(options, processLogger);
        var isolation = new IsolationProtocol(isolationLogger);
        _sessionManager = new SessionManager(
            options,
            _staThread,
            isolation,
            _processManager,
            sessionLogger
        );

        _hangDetectionTask = RunHangDetectionLoopAsync(_cts.Token);
        _idleShutdownTask =
            _options.IdleShutdownAfter > TimeSpan.Zero
                ? RunIdleShutdownLoopAsync(_cts.Token)
                : Task.CompletedTask;
    }

    public Task<TResult> ExecuteAsync<TResult>(
        Func<ISldWorks, TResult> operation,
        CancellationToken ct
    )
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, this);

        return _staThread.EnqueueAsync(
            () =>
            {
                // Lazy launch: ensure SW is running before first call
                _processManager.EnsureRunning(CancellationToken.None);

                var swApp =
                    _processManager.SwApp
                    ?? throw new InvalidOperationException("SolidWorks is not running");
                return operation(swApp);
            },
            ct
        );
    }

    public Task ExecuteAsync(Action<ISldWorks> operation, CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, this);

        return _staThread.EnqueueAsync(
            () =>
            {
                _processManager.EnsureRunning(CancellationToken.None);

                var swApp =
                    _processManager.SwApp
                    ?? throw new InvalidOperationException("SolidWorks is not running");
                operation(swApp);
            },
            ct
        );
    }

    public Task<ISwSession> AcquireSessionAsync(
        string workingDirectory,
        TimeSpan timeout,
        CancellationToken ct
    )
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, this);

        return _sessionManager
            .AcquireAsync(workingDirectory, timeout, ct)
            .ContinueWith(t => (ISwSession)t.Result, TaskContinuationOptions.OnlyOnRanToCompletion);
    }

    public void RequestRestart(string reason)
    {
        _logger.LogWarning("Restart requested by caller: {Reason}", reason);
        _processManager.MarkDegraded();
        _processManager.Kill();
    }

    public ResourcePressure GetResourcePressure() => _processManager.GetResourcePressure();

    public SwWatchdogStatus GetStatus()
    {
        var snapshot = _processManager.SampleResources();

        return new SwWatchdogStatus
        {
            SwRunning = _processManager.IsRunning,
            SwPid = _processManager.SwPid,
            SwUptime = _processManager.Uptime,
            DocumentsOpen = GetDocumentCount(),
            ActiveSessionId = _sessionManager.ActiveSessionId,
            MemoryMb = _processManager.MemoryMb,
            Degraded = _processManager.IsDegraded,
            GdiObjects = snapshot?.GdiObjects ?? 0,
            UserObjects = snapshot?.UserObjects ?? 0,
            GdiPeak = snapshot?.GdiPeak ?? 0,
            FreeSystemMemoryMb = snapshot?.FreeSystemMemoryMb ?? 0,
            ResourcePressure = snapshot?.Overall ?? ResourcePressure.Low,
        };
    }

    private int GetDocumentCount()
    {
        if (!_processManager.IsRunning)
            return 0;
        try
        {
            return _staThread
                .EnqueueAsync(
                    () =>
                    {
                        var swApp = _processManager.SwApp;
                        return swApp?.GetDocumentCount() ?? 0;
                    },
                    CancellationToken.None
                )
                .GetAwaiter()
                .GetResult();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to query document count from SolidWorks");
            return -1;
        }
    }

    private async Task RunHangDetectionLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(_options.HangCheckInterval, ct);

                if (!_processManager.IsRunning)
                    continue;

                if (!_processManager.CheckResponsive())
                {
                    // Unresponsive — confirm with exponential backoff + CPU activity check.
                    // SW doesn't pump messages during heavy COM operations (large assembly load),
                    // so WM_NULL timeout alone causes false positives.
                    // CPU delta distinguishes busy (CPU growing) from hung (CPU stalled).
                    var retries = Math.Max(1, _options.HangConfirmRetries);
                    var cpuBefore = _processManager.GetCpuTime();
                    var confirmed = true;

                    for (var attempt = 1; attempt < retries; attempt++)
                    {
                        var backoff = TimeSpan.FromTicks(
                            _options.HangCheckInterval.Ticks * (1L << attempt)
                        );

                        await Task.Delay(backoff, ct);

                        if (_processManager.CheckResponsive())
                        {
                            _logger.LogInformation(
                                "SolidWorks recovered after {Attempt} retries (PID={Pid})",
                                attempt,
                                _processManager.SwPid
                            );
                            confirmed = false;
                            break;
                        }

                        // Check CPU activity — if CPU is growing, SW is busy, not hung
                        var cpuAfter = _processManager.GetCpuTime();
                        var cpuDelta = cpuAfter - cpuBefore;

                        if (cpuDelta > TimeSpan.FromMilliseconds(100))
                        {
                            _logger.LogInformation(
                                "SolidWorks unresponsive but CPU active (PID={Pid}, cpuDelta={CpuDelta}ms, attempt={Attempt}/{Retries}) — busy, not hung",
                                _processManager.SwPid,
                                cpuDelta.TotalMilliseconds,
                                attempt,
                                retries
                            );
                            confirmed = false;
                            break;
                        }

                        _logger.LogWarning(
                            "SolidWorks unresponsive, CPU idle (PID={Pid}, cpuDelta={CpuDelta}ms, attempt={Attempt}/{Retries})",
                            _processManager.SwPid,
                            cpuDelta.TotalMilliseconds,
                            attempt,
                            retries
                        );
                        cpuBefore = cpuAfter;
                    }

                    if (confirmed)
                    {
                        _logger.LogError(
                            "SolidWorks hang confirmed — unresponsive and CPU idle after {Retries} checks (PID={Pid}), killing",
                            retries,
                            _processManager.SwPid
                        );

                        // Kill — process is truly hung, nothing to save.
                        // If a session is active, the blocked COM call will throw COMException
                        // which SwSession wraps into SwProcessKilledException for the caller.
                        // Next AcquireSessionAsync will restart SW via EnsureRunning.
                        _processManager.MarkDegraded();
                        _processManager.Kill();
                    }
                }
                else if (_processManager.CheckModalDialog())
                {
                    // D7 Layer 5: Process responds to WM_NULL but main window is disabled
                    // — a modal dialog is blocking COM calls.
                    _logger.LogWarning(
                        "Modal dialog detected on SolidWorks (PID={Pid}), main window disabled",
                        _processManager.SwPid
                    );

                    // D7 Layer 6: Try auto-dismiss via WM_CLOSE
                    var dismissed = _processManager.TryDismissModalDialogs();
                    if (dismissed > 0)
                    {
                        _logger.LogInformation(
                            "D7 Layer 6: Sent WM_CLOSE to {Count} dialog(s), waiting for effect",
                            dismissed
                        );

                        // Wait briefly for WM_CLOSE to take effect, then re-check
                        await Task.Delay(TimeSpan.FromMilliseconds(500), ct);

                        if (!_processManager.CheckModalDialog())
                        {
                            _logger.LogInformation(
                                "D7 Layer 6: Modal dialog dismissed successfully (PID={Pid})",
                                _processManager.SwPid
                            );
                            continue; // Dialog cleared — back to normal monitoring
                        }

                        _logger.LogWarning(
                            "D7 Layer 6: WM_CLOSE did not clear modal dialog (PID={Pid}), escalating to Layer 7",
                            _processManager.SwPid
                        );
                    }

                    // D7 Layer 7: Kill as last resort
                    _logger.LogError(
                        "Modal dialog could not be auto-dismissed (PID={Pid}) — killing",
                        _processManager.SwPid
                    );
                    _processManager.MarkDegraded();
                    _processManager.Kill();
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Hang detection loop error");
            }
        }
    }

    private async Task RunIdleShutdownLoopAsync(CancellationToken ct)
    {
        var lastActivity = DateTime.UtcNow;

        while (!ct.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromMinutes(1), ct);

                if (!_processManager.IsRunning)
                    continue;
                if (_sessionManager.ActiveSessionId is not null)
                {
                    lastActivity = DateTime.UtcNow;
                    continue;
                }

                if (DateTime.UtcNow - lastActivity > _options.IdleShutdownAfter)
                {
                    _logger.LogInformation(
                        "Idle shutdown: no activity for {Idle}",
                        _options.IdleShutdownAfter
                    );
                    _processManager.Kill();
                    lastActivity = DateTime.UtcNow; // Reset for next launch
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Idle shutdown loop error");
            }
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;

        await _cts.CancelAsync();

        try
        {
            await _hangDetectionTask;
        }
        catch (OperationCanceledException) { }
        try
        {
            await _idleShutdownTask;
        }
        catch (OperationCanceledException) { }

        _sessionManager.Dispose();
        _processManager.Dispose();
        _staThread.Dispose();
        _cts.Dispose();

        _logger.LogInformation("SwWatchdog disposed");
    }
}
