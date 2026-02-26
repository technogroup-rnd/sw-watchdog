using System.Runtime.InteropServices;

namespace SwWatchdog.Sta;

/// <summary>
/// COM IMessageFilter implementation for handling RPC_E_SERVERCALL_RETRYLATER
/// from SolidWorks during startup and busy periods.
/// Must be registered on the STA thread via CoRegisterMessageFilter.
/// </summary>
[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("00000016-0000-0000-C000-000000000046")]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, nint hTaskCaller, int dwTickCount, nint lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(nint hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(nint hTaskCallee, int dwTickCount, int dwPendingType);
}

internal sealed class ComMessageFilter : IOleMessageFilter
{
    private const int SERVERCALL_ISHANDLED = 0;
    private const int PENDINGMSG_WAITDEFPROCESS = 2;
    private const int SERVERCALL_RETRYLATER = 2;

    private readonly int _maxRetryMs;

    public ComMessageFilter(int maxRetryMs = 10_000)
    {
        _maxRetryMs = maxRetryMs;
    }

    public int HandleInComingCall(
        int dwCallType,
        nint hTaskCaller,
        int dwTickCount,
        nint lpInterfaceInfo
    )
    {
        return SERVERCALL_ISHANDLED;
    }

    public int RetryRejectedCall(nint hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == SERVERCALL_RETRYLATER && dwTickCount < _maxRetryMs)
            return 100; // retry after 100ms

        return -1; // give up — will throw COMException
    }

    public int MessagePending(nint hTaskCallee, int dwTickCount, int dwPendingType)
    {
        return PENDINGMSG_WAITDEFPROCESS;
    }

    // --- Registration ---

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IOleMessageFilter? lpMessageFilter,
        out IOleMessageFilter? lplpMessageFilter
    );

    private static IOleMessageFilter? _previousFilter;

    /// <summary>
    /// Register this filter on the current STA thread.
    /// Must be called BEFORE any COM interaction (even ROT polling).
    /// </summary>
    public static void Register(int maxRetryMs = 10_000)
    {
        CoRegisterMessageFilter(new ComMessageFilter(maxRetryMs), out _previousFilter);
    }

    /// <summary>
    /// Restore the previous filter (or null). Call on STA thread shutdown.
    /// </summary>
    public static void Unregister()
    {
        CoRegisterMessageFilter(_previousFilter, out _);
        _previousFilter = null;
    }
}
