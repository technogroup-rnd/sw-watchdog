using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

namespace SwWatchdog.Process;

/// <summary>
/// Samples GDI/USER objects and system memory for a process.
/// All methods are Win32-only (no COM, no STA) — safe to call from any thread.
/// </summary>
internal sealed class ResourceMonitor
{
    private readonly SwWatchdogOptions _options;
    private readonly ILogger _logger;
    private readonly uint _gdiLimit;
    private readonly uint _userLimit;

    public ResourceMonitor(SwWatchdogOptions options, ILogger logger)
    {
        _options = options;
        _logger = logger;
        (_gdiLimit, _userLimit) = ReadLimitsFromRegistry();
        _logger.LogInformation(
            "Resource limits from registry: GDI={GdiLimit}, USER={UserLimit}",
            _gdiLimit,
            _userLimit
        );
    }

    /// <summary>
    /// Take a full resource snapshot for the given process.
    /// Throws <see cref="Win32Exception"/> or <see cref="ArgumentException"/> on failure.
    /// </summary>
    public ResourceSnapshot Sample(int pid)
    {
        // Process.Handle requires the object not to be disposed — keep inside using.
        using var proc = System.Diagnostics.Process.GetProcessById(pid);
        var handle = proc.Handle;

        var gdi = QueryGuiResource(handle, NativeMethods.GR_GDIOBJECTS, pid, "GDI");
        var user = QueryGuiResource(handle, NativeMethods.GR_USEROBJECTS, pid, "USER");
        var gdiPeak = QueryGuiResource(handle, NativeMethods.GR_GDIOBJECTS_PEAK, pid, "GDI_PEAK");

        var freeMemoryMb = GetFreeSystemMemoryMb();

        var gdiThreshold = _gdiLimit * _options.GdiObjectsLimitPercent / 100.0;
        var userThreshold = _userLimit * _options.UserObjectsLimitPercent / 100.0;

        var gdiPressure =
            gdiThreshold > 0
                ? ResourcePressureCalculator.FromRatio(gdi / gdiThreshold)
                : ResourcePressure.Low;
        var userPressure =
            userThreshold > 0
                ? ResourcePressureCalculator.FromRatio(user / userThreshold)
                : ResourcePressure.Low;
        // Memory: inverted ratio — limit/actual, so low free memory → ratio > 1.0 → Critical
        var memoryPressure =
            _options.FreeMemoryLimitMb > 0 && freeMemoryMb > 0
                ? ResourcePressureCalculator.FromRatio(
                    _options.FreeMemoryLimitMb / (double)freeMemoryMb
                )
                : ResourcePressure.Low;

        var overall = (ResourcePressure)
            Math.Max(Math.Max((int)gdiPressure, (int)userPressure), (int)memoryPressure);

        return new ResourceSnapshot
        {
            GdiObjects = gdi,
            UserObjects = user,
            GdiPeak = gdiPeak,
            FreeSystemMemoryMb = freeMemoryMb,
            GdiPressure = gdiPressure,
            UserPressure = userPressure,
            MemoryPressure = memoryPressure,
            Overall = overall,
        };
    }

    /// <summary>
    /// Returns the worst resource pressure level. On any error, returns <see cref="ResourcePressure.Low"/>.
    /// </summary>
    public ResourcePressure GetPressure(int pid)
    {
        try
        {
            return Sample(pid).Overall;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogWarning(ex, "Failed to sample resource pressure for PID={Pid}", pid);
            return ResourcePressure.Low;
        }
    }

    private static uint QueryGuiResource(nint handle, uint flag, int pid, string name)
    {
        var value = NativeMethods.GetGuiResources(handle, flag);
        if (value == 0)
        {
            var err = Marshal.GetLastWin32Error();
            if (err != 0)
                throw new Win32Exception(err, $"GetGuiResources({name}) failed for PID={pid}");
            // value==0 && err==0: process has no GUI objects — unlikely for SW but not an error
        }
        return value;
    }

    private static long GetFreeSystemMemoryMb()
    {
        var status = new NativeMethods.MEMORYSTATUSEX
        {
            dwLength = (uint)Marshal.SizeOf<NativeMethods.MEMORYSTATUSEX>(),
        };

        if (!NativeMethods.GlobalMemoryStatusEx(ref status))
        {
            var err = Marshal.GetLastWin32Error();
            throw new Win32Exception(err, "GlobalMemoryStatusEx failed");
        }

        return (long)(status.ullAvailPhys / (1024 * 1024));
    }

    /// <summary>
    /// Read GDI/USER process handle quotas from the Windows registry.
    /// Returns 10000 (system default) if keys are not found.
    /// </summary>
    private static (uint Gdi, uint User) ReadLimitsFromRegistry()
    {
        const string keyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows";
        using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath);
        if (key is null)
            return (10_000, 10_000);

        var gdi = Convert.ToUInt32(key.GetValue("GDIProcessHandleQuota", 10_000));
        var user = Convert.ToUInt32(key.GetValue("USERProcessHandleQuota", 10_000));
        return (gdi, user);
    }
}

/// <summary>
/// Point-in-time snapshot of GDI, USER, and system memory resources.
/// </summary>
internal sealed record ResourceSnapshot
{
    public required uint GdiObjects { get; init; }
    public required uint UserObjects { get; init; }
    public required uint GdiPeak { get; init; }
    public required long FreeSystemMemoryMb { get; init; }
    public required ResourcePressure GdiPressure { get; init; }
    public required ResourcePressure UserPressure { get; init; }
    public required ResourcePressure MemoryPressure { get; init; }
    public required ResourcePressure Overall { get; init; }
}
