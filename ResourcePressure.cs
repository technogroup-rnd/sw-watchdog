namespace SwWatchdog;

/// <summary>
/// Resource pressure level across GDI, USER, and system memory.
/// Pressure ratio = resource_used / configured_restart_threshold.
/// 100% (Critical) = at the configured restart threshold, not the absolute system limit.
/// </summary>
public enum ResourcePressure
{
    Low = 0, // ratio < 0.50
    Moderate = 1, // ratio 0.50–0.66
    Elevated = 2, // ratio 0.67–0.74
    High = 3, // ratio 0.75–0.99
    Critical = 4, // ratio >= 1.00
}

public static class ResourcePressureCalculator
{
    /// <summary>
    /// Convert a pressure ratio (0.0–1.0+) to a <see cref="ResourcePressure"/> level.
    /// </summary>
    public static ResourcePressure FromRatio(double ratio) =>
        ratio switch
        {
            >= 1.00 => ResourcePressure.Critical,
            >= 0.75 => ResourcePressure.High,
            >= 0.67 => ResourcePressure.Elevated,
            >= 0.50 => ResourcePressure.Moderate,
            _ => ResourcePressure.Low,
        };
}
