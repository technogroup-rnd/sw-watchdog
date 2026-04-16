namespace SwWatchdog;

/// <summary>
/// Current state of the watchdog and managed SolidWorks process.
/// </summary>
public sealed record SwWatchdogStatus
{
    public required bool SwRunning { get; init; }
    public required int SwPid { get; init; }
    public required TimeSpan SwUptime { get; init; }
    public required string? ActiveSessionId { get; init; }
    public required long MemoryMb { get; init; }
    public required bool Degraded { get; init; }
    public required uint GdiObjects { get; init; }
    public required uint UserObjects { get; init; }
    public required uint GdiPeak { get; init; }
    public required long FreeSystemMemoryMb { get; init; }
    public required ResourcePressure ResourcePressure { get; init; }
}
