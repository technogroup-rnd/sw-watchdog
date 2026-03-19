namespace SwWatchdog;

/// <summary>
/// Current state of the watchdog and managed SolidWorks process.
/// </summary>
public sealed record SwWatchdogStatus
{
    public required bool SwRunning { get; init; }
    public required int SwPid { get; init; }
    public required TimeSpan SwUptime { get; init; }
    public required int DocumentsOpen { get; init; }
    public required string? ActiveSessionId { get; init; }
    public required long MemoryMb { get; init; }
    public required bool Degraded { get; init; }
}
