using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwWatchdog.Process;

/// <summary>
/// Manages the SolidWorks process lifecycle: 3-stage launch, ROT detection,
/// startup completion, hang detection, crash detection, restart, idle shutdown.
/// All public methods are thread-safe. COM calls must be invoked on the STA thread.
/// </summary>
internal sealed class SwProcessManager : IDisposable
{
    private readonly SwWatchdogOptions _options;
    private readonly ILogger _logger;
    private readonly Lock _lock = new();

    private System.Diagnostics.Process? _swProcess;
    private ISldWorks? _swApp;
    private nint _swHwnd;
    private DateTime _startTime;
    private int _filesProcessed;
    private bool _tainted;
    private bool _disposed;

    /// <summary>
    /// Fired when SW process crashes unexpectedly.
    /// </summary>
    public event Action<int>? ProcessCrashed;

    public SwProcessManager(IOptions<SwWatchdogOptions> options, ILogger logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public ISldWorks? SwApp
    {
        get
        {
            lock (_lock)
            {
                return _swApp;
            }
        }
    }

    public bool IsRunning
    {
        get
        {
            lock (_lock)
            {
                return _swProcess is { HasExited: false } && _swApp is not null;
            }
        }
    }

    public bool IsTainted
    {
        get
        {
            lock (_lock)
            {
                return _tainted;
            }
        }
    }

    public int SwPid
    {
        get
        {
            lock (_lock)
            {
                return _swProcess?.Id ?? 0;
            }
        }
    }

    public TimeSpan Uptime
    {
        get
        {
            lock (_lock)
            {
                return _swProcess is { HasExited: false }
                    ? DateTime.UtcNow - _startTime
                    : TimeSpan.Zero;
            }
        }
    }

    public long MemoryMb
    {
        get
        {
            lock (_lock)
            {
                if (_swProcess is not { HasExited: false })
                    return 0;
                try
                {
                    _swProcess.Refresh();
                    return _swProcess.WorkingSet64 / (1024 * 1024);
                }
                catch
                {
                    return 0;
                }
            }
        }
    }

    public void MarkTainted()
    {
        lock (_lock)
        {
            _tainted = true;
        }
        _logger.LogWarning("SolidWorks process marked as tainted");
    }

    public void IncrementFileCount(int count = 1)
    {
        lock (_lock)
        {
            _filesProcessed += count;
        }
    }

    /// <summary>
    /// Ensure SolidWorks is running. If not, launch and wait for full readiness.
    /// MUST be called from the STA thread.
    /// </summary>
    public void EnsureRunning(CancellationToken ct)
    {
        if (IsRunning)
            return;
        Launch(ct);
    }

    /// <summary>
    /// Check if periodic restart is needed (between sessions, not during).
    /// </summary>
    public bool NeedsRestart()
    {
        lock (_lock)
        {
            if (_tainted)
                return true;
            if (
                _options.RestartAfterFileCount > 0
                && _filesProcessed >= _options.RestartAfterFileCount
            )
                return true;
            if (
                _options.RestartAfterElapsed > TimeSpan.Zero
                && Uptime >= _options.RestartAfterElapsed
            )
                return true;
            if (_options.RestartAfterMemoryMb > 0 && MemoryMb >= _options.RestartAfterMemoryMb)
                return true;
            return false;
        }
    }

    /// <summary>
    /// Kill and restart SolidWorks. MUST be called from the STA thread.
    /// </summary>
    public void Restart(CancellationToken ct)
    {
        _logger.LogInformation(
            "Restarting SolidWorks (tainted={Tainted}, files={Files}, uptime={Uptime}, memMb={Mem})",
            _tainted,
            _filesProcessed,
            Uptime,
            MemoryMb
        );

        Kill();
        Launch(ct);
    }

    /// <summary>
    /// Check if the SW window is responding (hang detection).
    /// Returns true if responding, false if hung.
    /// Can be called from any thread.
    /// </summary>
    public bool CheckResponsive()
    {
        nint hwnd;
        lock (_lock)
        {
            if (_swProcess is not { HasExited: false })
                return false;
            hwnd = _swHwnd;
        }

        if (hwnd == 0)
            return true; // No window yet = not hung

        var timeoutMs = (uint)_options.HangCheckTimeout.TotalMilliseconds;
        var result = NativeMethods.SendMessageTimeoutW(
            hwnd,
            NativeMethods.WM_NULL,
            0,
            0,
            NativeMethods.SMTO_ABORTIFHUNG,
            timeoutMs,
            out _
        );

        return result != 0;
    }

    /// <summary>
    /// Check if a modal dialog is blocking the SW main window (D7, Layer 5).
    /// Returns true if a modal dialog is detected, false otherwise.
    /// Uses IsWindowEnabled — disabled main window means a modal child exists.
    /// Can be called from any thread.
    /// </summary>
    public bool CheckModalDialog()
    {
        nint hwnd;
        lock (_lock)
        {
            if (_swProcess is not { HasExited: false })
                return false;
            hwnd = _swHwnd;
        }

        if (hwnd == 0)
            return false; // No window yet = can't check

        return !NativeMethods.IsWindowEnabled(hwnd);
    }

    /// <summary>
    /// Attempt to auto-dismiss modal dialogs on the SW process (D7, Layer 6).
    /// Finds visible #32770 dialog windows owned by the SW process via EnumWindows,
    /// then sends WM_CLOSE to each.
    /// IMPORTANT: Uses EnumWindows + PID filter (not EnumThreadWindows) because
    /// SW creates dialogs on different threads than the main window.
    /// Returns the number of dialogs WM_CLOSE was sent to.
    /// Can be called from any thread.
    /// </summary>
    public int TryDismissModalDialogs()
    {
        int pid;
        lock (_lock)
        {
            if (_swProcess is not { HasExited: false })
                return 0;
            pid = _swProcess.Id;
        }

        var dialogHandles = new List<nint>();

        NativeMethods.EnumWindows((hwnd, _) =>
        {
            NativeMethods.GetWindowThreadProcessId(hwnd, out var windowPid);
            if (windowPid != (uint)pid)
                return true; // different process, continue

            var className = NativeMethods.GetClassName(hwnd);
            if (className == "#32770" && NativeMethods.IsWindowVisible(hwnd))
                dialogHandles.Add(hwnd);

            return true; // continue enumeration
        }, nint.Zero);

        foreach (var dlg in dialogHandles)
        {
            _logger.LogWarning(
                "D7 Layer 6: Sending WM_CLOSE to dialog 0x{Handle:X} (PID={Pid})",
                dlg, pid);
            NativeMethods.PostMessage(dlg, NativeMethods.WM_CLOSE, nint.Zero, nint.Zero);
        }

        return dialogHandles.Count;
    }

    /// <summary>
    /// Kill the SolidWorks process.
    /// </summary>
    public void Kill()
    {
        lock (_lock)
        {
            if (_swApp is not null)
            {
                try
                {
                    Marshal.ReleaseComObject(_swApp);
                }
                catch { }
                _swApp = null;
            }

            _swHwnd = 0;

            if (_swProcess is not null)
            {
                _swProcess.Exited -= OnProcessExited;

                if (!_swProcess.HasExited)
                {
                    try
                    {
                        _logger.LogInformation("Killing SolidWorks (PID={Pid})", _swProcess.Id);
                        _swProcess.Kill(entireProcessTree: true);
                        _swProcess.WaitForExit(TimeSpan.FromSeconds(10));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to kill SolidWorks process");
                    }
                }

                _swProcess.Dispose();
                _swProcess = null;
            }

            _tainted = false;
            _filesProcessed = 0;
        }
    }

    /// <summary>
    /// Full 3-stage launch: Process.Start(/r) → ROT polling → OnIdleNotify + StartupProcessCompleted.
    /// MUST be called from the STA thread (for COM event subscription).
    /// </summary>
    private void Launch(CancellationToken ct)
    {
        var exePath = ResolveSolidWorksPath();
        _logger.LogInformation("Stage 1: Launching SolidWorks from {Path}", exePath);

        // --- Stage 1: Process.Start ---
        System.Diagnostics.Process process;
        lock (_lock)
        {
            _swProcess?.Dispose();
            process = System.Diagnostics.Process.Start(
                new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = "/r", // suppress splash screen
                    UseShellExecute = true,
                }
            )!;
            _swProcess = process;
            _startTime = DateTime.UtcNow;
            _tainted = false;
            _filesProcessed = 0;

            // Crash detection via event (instant, not polling)
            process.EnableRaisingEvents = true;
            process.Exited += OnProcessExited;
        }

        _logger.LogInformation("SolidWorks started (PID={Pid}), waiting for ROT...", process.Id);

        var deadline = DateTime.UtcNow + _options.StartupTimeout;

        // --- Stage 2: ROT polling ---
        var swApp = WaitForRotRegistration(process.Id, deadline, ct);
        _logger.LogInformation(
            "Stage 2 complete: ISldWorks obtained from ROT (PID={Pid})",
            process.Id
        );

        // --- Stage 3: Wait for full startup ---
        WaitForStartupCompletion(swApp, process, deadline, ct);
        _logger.LogInformation(
            "Stage 3 complete: StartupProcessCompleted=true (PID={Pid})",
            process.Id
        );

        // --- Acquire HWND via COM (not Process.MainWindowHandle — .NET bug #32690) ---
        nint hwnd = 0;
        try
        {
            var frame = (IFrame?)swApp.Frame();
            if (frame is not null)
                hwnd = new nint(frame.GetHWnd());
        }
        catch (Exception ex)
        {
            _logger.LogWarning(
                ex,
                "Failed to get HWND via IFrame — hang detection will use fallback"
            );
        }

        // --- Apply performance settings ---
        try
        {
            swApp.CommandInProgress = true;
            swApp.FrameState = (int)swWindowState_e.swWindowMinimized;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to set performance options");
        }

        // --- Apply dialog suppression preferences (D7, Layer 2) ---
        try
        {
            swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swShowErrorsEveryRebuild, false);
            swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swExtRefNoPromptOrSave, true);
            swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swWhileOpeningAssembliesAutoDismissMessages, true);
            swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swWarnSaveUpdateErrors, false);
            swApp.SetUserPreferenceToggle(
                (int)swUserPreferenceToggle_e.swWarnSavingReferencedDoc, false);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to set dialog suppression preferences");
        }

        lock (_lock)
        {
            _swApp = swApp;
            _swHwnd = hwnd;
        }

        _logger.LogInformation(
            "SolidWorks fully ready (PID={Pid}, HWND=0x{Hwnd:X})",
            process.Id,
            hwnd
        );
    }

    private ISldWorks WaitForRotRegistration(int pid, DateTime deadline, CancellationToken ct)
    {
        var rotMoniker = $"SolidWorks_PID_{pid}";

        while (DateTime.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();

            lock (_lock)
            {
                if (_swProcess is { HasExited: true })
                    throw new InvalidOperationException(
                        $"SolidWorks process exited during startup (exit code: {_swProcess.ExitCode})"
                    );
            }

            var swApp = TryGetFromRot(rotMoniker);
            if (swApp is not null)
                return swApp;

            Thread.Sleep(_options.RotPollingInterval);
        }

        throw new TimeoutException(
            $"SolidWorks (PID={pid}) did not register in ROT within {_options.StartupTimeout}"
        );
    }

    /// <summary>
    /// Wait for StartupProcessCompleted using OnIdleNotify COM event.
    /// Requires STA thread with message pump (WinForms).
    /// </summary>
    private void WaitForStartupCompletion(
        ISldWorks swApp,
        System.Diagnostics.Process process,
        DateTime deadline,
        CancellationToken ct
    )
    {
        // Quick check — might already be complete
        try
        {
            if (swApp.StartupProcessCompleted)
                return;
        }
        catch (COMException)
        {
            // May throw during early startup — continue to event-based wait
        }

        bool isReady = false;

        var onIdleHandler = new DSldWorksEvents_OnIdleNotifyEventHandler(() =>
        {
            try
            {
                if (swApp.StartupProcessCompleted)
                    isReady = true;
            }
            catch (COMException)
            {
                // Not ready yet — will retry on next idle
            }
            return 0;
        });

        try
        {
            // Subscribe via SldWorks (not ISldWorks) — events are on the class, not interface
            ((SldWorks)swApp).OnIdleNotify += onIdleHandler;

            while (!isReady)
            {
                ct.ThrowIfCancellationRequested();

                if (DateTime.UtcNow > deadline)
                    throw new TimeoutException(
                        $"SolidWorks (PID={process.Id}) did not complete startup within {_options.StartupTimeout}"
                    );

                if (process.HasExited)
                    throw new InvalidOperationException(
                        $"SolidWorks process exited during startup (exit code: {process.ExitCode})"
                    );

                // Pump messages to receive COM events — DoEvents returns immediately
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(50);
            }
        }
        finally
        {
            ((SldWorks)swApp).OnIdleNotify -= onIdleHandler;
        }
    }

    private void OnProcessExited(object? sender, EventArgs e)
    {
        int exitCode;
        lock (_lock)
        {
            if (_swProcess is null)
                return;
            exitCode = _swProcess.ExitCode;
            _swApp = null;
            _swHwnd = 0;
        }

        _logger.LogError("SolidWorks process crashed (ExitCode={ExitCode})", exitCode);
        ProcessCrashed?.Invoke(exitCode);
    }

    private static ISldWorks? TryGetFromRot(string monikerName)
    {
        if (NativeMethods.GetRunningObjectTable(0, out var rot) != 0)
            return null;

        try
        {
            rot.EnumRunning(out var enumMoniker);
            var monikers = new IMoniker[1];

            while (enumMoniker.Next(1, monikers, nint.Zero) == 0)
            {
                if (NativeMethods.CreateBindCtx(0, out var ctx) != 0)
                    continue;

                try
                {
                    monikers[0].GetDisplayName(ctx, null!, out var displayName);

                    if (string.Equals(displayName, monikerName, StringComparison.OrdinalIgnoreCase))
                    {
                        rot.GetObject(monikers[0], out var obj);
                        return obj as ISldWorks;
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    // Moniker from another user/privilege level — skip
                }
                finally
                {
                    Marshal.ReleaseComObject(ctx);
                }
            }

            return null;
        }
        finally
        {
            Marshal.ReleaseComObject(rot);
        }
    }

    private string ResolveSolidWorksPath()
    {
        if (!string.IsNullOrEmpty(_options.SolidWorksPath))
        {
            if (!File.Exists(_options.SolidWorksPath))
                throw new FileNotFoundException(
                    "Configured SolidWorks path not found",
                    _options.SolidWorksPath
                );
            return _options.SolidWorksPath;
        }

        // Try standard registry location
        var regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
            @"SOFTWARE\SolidWorks\SolidWorks\"
        );

        if (regKey is not null)
        {
            var installPath = regKey.GetValue("SolidWorksPath") as string;
            if (!string.IsNullOrEmpty(installPath))
            {
                var exePath = Path.Combine(installPath, "SLDWORKS.exe");
                if (File.Exists(exePath))
                    return exePath;
            }
        }

        // Fallback: common install path
        var defaultPath = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe";
        if (File.Exists(defaultPath))
            return defaultPath;

        throw new FileNotFoundException(
            "SolidWorks installation not found. Set SwWatchdogOptions.SolidWorksPath explicitly."
        );
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        Kill();
    }
}
