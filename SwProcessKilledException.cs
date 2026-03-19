namespace SwWatchdog;

/// <summary>
/// Thrown when a COM operation fails because SolidWorks was killed by hang detection.
/// Callers should treat this as a transient error — retry with a fresh SW session.
/// Unlike a raw COMException (which may indicate a corrupt file or model error),
/// this exception means the operation was interrupted and may succeed on retry.
/// </summary>
public sealed class SwProcessKilledException : Exception
{
    public SwProcessKilledException(string message, Exception innerException)
        : base(message, innerException) { }
}
