namespace SwWatchdog;

/// <summary>
/// Thrown when a COM operation fails because SolidWorks was killed by the watchdog
/// (hang detection or modal dialog escalation) during an active session.
/// Callers should treat this as a transient error — retry with a fresh SW session.
/// Unlike a raw COMException (which may indicate a corrupt file or model error),
/// this exception means the operation was interrupted and may succeed on retry.
/// The gRPC layer maps this to <c>StatusCode.Unavailable</c> (conventional retry signal).
/// </summary>
public sealed class SwProcessKilledException : Exception
{
    public SwProcessKilledException(string message, Exception innerException)
        : base(message, innerException) { }
}
