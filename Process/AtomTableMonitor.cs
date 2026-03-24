using System.Text;
using Microsoft.Extensions.Logging;

namespace SwWatchdog.Process;

/// <summary>
/// Monitors Global Atom Table usage — a per-WindowStation resource that accumulates
/// through RegisterWindowMessage calls in SolidWorks. Atoms have no Unregister API:
/// killing the process does NOT free them. Only a machine reboot clears the table.
/// <para>
/// Full enumeration scans all 16384 slots (0xC000..0xFFFF) via GlobalGetAtomName.
/// Takes ~25-40 ms — call periodically (every N file ops), not on every operation.
/// </para>
/// <para>
/// This is the Level 2 (session pressure) monitor. For Level 1 (process pressure:
/// GDI/USER/Memory), see <see cref="ResourceMonitor"/>.
/// </para>
/// </summary>
public sealed class AtomTableMonitor
{
    private readonly int _thresholdPercent;
    private readonly ILogger _logger;

    /// <param name="thresholdPercent">
    /// Atom table usage percentage (0-100) at which pressure reaches Critical.
    /// Mapped from <c>RecycleOptions.AtomTablePercentThreshold</c> by the consumer.
    /// 0 = monitoring disabled (Sample always returns Low pressure).
    /// </param>
    /// <param name="logger">Logger for diagnostic output after each enumeration.</param>
    public AtomTableMonitor(int thresholdPercent, ILogger logger)
    {
        _thresholdPercent = thresholdPercent;
        _logger = logger;
    }

    /// <summary>
    /// Full enumeration of the Global Atom Table. Counts occupied slots and calculates
    /// pressure relative to the configured threshold.
    /// ~25-40 ms on typical hardware. Safe to call from any thread (Win32 only, no COM/STA).
    /// </summary>
    public AtomTableSnapshot Sample()
    {
        var used = CountUsedAtoms();
        var snapshot = AtomTableSnapshot.Create(used, _thresholdPercent);

        _logger.LogDebug(
            "Atom table: {UsedAtoms}/{MaxAtoms} ({UsagePercent:F1}%), pressure={Pressure}",
            snapshot.UsedAtoms,
            AtomTableSnapshot.MaxAtoms,
            snapshot.UsagePercent,
            snapshot.Pressure
        );

        return snapshot;
    }

    /// <summary>
    /// Count occupied atoms in the Global Atom Table by scanning the full range 0xC000..0xFFFF.
    /// Each slot is tested with GlobalGetAtomName — returns >0 if the atom exists.
    /// </summary>
    public static int CountUsedAtoms()
    {
        var used = 0;
        var buffer = new StringBuilder(256);

        for (var atom = AtomTableSnapshot.AtomRangeStart;
             atom <= AtomTableSnapshot.AtomRangeEnd;
             atom++)
        {
            buffer.Clear();
            if (NativeMethods.GlobalGetAtomName((ushort)atom, buffer, buffer.Capacity) > 0)
                used++;
        }

        return used;
    }
}

/// <summary>
/// Point-in-time snapshot of Global Atom Table usage with pressure level.
/// Pressure is relative to the configured threshold (not the absolute system maximum):
/// 100% (Critical) = at the configured reboot threshold, not at 16384 atoms.
/// </summary>
public sealed record AtomTableSnapshot
{
    /// <summary>Start of the atom range (0xC000 — first application atom).</summary>
    public const int AtomRangeStart = 0xC000;

    /// <summary>End of the atom range (0xFFFF — last possible atom).</summary>
    public const int AtomRangeEnd = 0xFFFF;

    /// <summary>Total capacity of the Global Atom Table (16384 slots).</summary>
    public const int MaxAtoms = AtomRangeEnd - AtomRangeStart + 1;

    /// <summary>Number of atoms currently in use.</summary>
    public required int UsedAtoms { get; init; }

    /// <summary>
    /// Usage as percentage of <see cref="MaxAtoms"/> (0.0–100.0).
    /// This is the absolute fill level, independent of the configured threshold.
    /// </summary>
    public double UsagePercent => UsedAtoms * 100.0 / MaxAtoms;

    /// <summary>
    /// Pressure level relative to the configured threshold (not the absolute max).
    /// E.g., with threshold=60%: usage 45% → High (75% of threshold), usage 60% → Critical.
    /// </summary>
    public required ResourcePressure Pressure { get; init; }

    /// <summary>
    /// Create a snapshot with pressure calculated from the configured threshold.
    /// </summary>
    /// <param name="usedAtoms">Number of atoms currently in use (from enumeration).</param>
    /// <param name="thresholdPercent">
    /// The percentage of MaxAtoms at which pressure reaches Critical.
    /// 0 = disabled (always returns Low).
    /// </param>
    public static AtomTableSnapshot Create(int usedAtoms, int thresholdPercent)
    {
        var usagePercent = usedAtoms * 100.0 / MaxAtoms;
        var ratio = thresholdPercent > 0 ? usagePercent / thresholdPercent : 0.0;

        return new AtomTableSnapshot
        {
            UsedAtoms = usedAtoms,
            Pressure = ResourcePressureCalculator.FromRatio(ratio),
        };
    }
}
