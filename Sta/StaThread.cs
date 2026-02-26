using System.Collections.Concurrent;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;

namespace SwWatchdog.Sta;

/// <summary>
/// Dedicated STA thread with WinForms message pump.
/// Supports COM events (OnIdleNotify, StartupProcessCompleted) and IMessageFilter.
/// Work items are dispatched via WindowsFormsSynchronizationContext.
/// </summary>
internal sealed class StaThread : IDisposable
{
    private readonly Thread _thread;
    private readonly ILogger _logger;
    private readonly ManualResetEventSlim _ready = new(false);
    private readonly ConcurrentQueue<Action> _startupQueue = new();

    private WindowsFormsSynchronizationContext? _syncCtx;
    private ApplicationContext? _appCtx;
    private bool _disposed;

    public StaThread(ILogger logger)
    {
        _logger = logger;

        _thread = new Thread(RunLoop) { Name = "SwWatchdog-STA", IsBackground = true };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();

        // Wait for the message pump to start before accepting work
        _ready.Wait();
    }

    /// <summary>
    /// Enqueue a function that returns a result. Runs on the STA thread.
    /// </summary>
    public Task<TResult> EnqueueAsync<TResult>(Func<TResult> work, CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var tcs = new TaskCompletionSource<TResult>(
            TaskCreationOptions.RunContinuationsAsynchronously
        );

        Post(() =>
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                tcs.TrySetResult(work());
            }
            catch (OperationCanceledException)
            {
                tcs.TrySetCanceled(ct);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });

        return tcs.Task;
    }

    /// <summary>
    /// Enqueue an action (no result). Runs on the STA thread.
    /// </summary>
    public Task EnqueueAsync(Action work, CancellationToken ct)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        Post(() =>
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                work();
                tcs.TrySetResult();
            }
            catch (OperationCanceledException)
            {
                tcs.TrySetCanceled(ct);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });

        return tcs.Task;
    }

    private void Post(Action action)
    {
        if (_syncCtx is not null)
        {
            _syncCtx.Post(_ => action(), null);
        }
        else
        {
            // SyncContext not yet available — queue for later
            _startupQueue.Enqueue(action);
        }
    }

    private void RunLoop()
    {
        _logger.LogInformation(
            "STA thread started (ThreadId={ThreadId})",
            Environment.CurrentManagedThreadId
        );

        try
        {
            // Register IMessageFilter BEFORE any COM interaction
            ComMessageFilter.Register();

            // Set up WinForms synchronization context for this STA thread
            WindowsFormsSynchronizationContext.AutoInstall = false;
            _syncCtx = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(_syncCtx);

            _appCtx = new ApplicationContext();

            // Signal that the pump is ready
            _ready.Set();

            // Drain anything that was queued before _syncCtx was available
            while (_startupQueue.TryDequeue(out var queued))
                queued();

            // Run the WinForms message loop — pumps Windows messages,
            // delivers COM events, processes SynchronizationContext.Post() callbacks
            Application.Run(_appCtx);
        }
        catch (Exception ex)
        {
            _logger.LogCritical(ex, "STA thread crashed");
        }
        finally
        {
            ComMessageFilter.Unregister();
        }

        _logger.LogInformation("STA thread exiting");
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;

        // Exit the Application.Run message loop from any thread
        if (_appCtx is not null)
        {
            // Post ExitThread to the STA message loop
            _syncCtx?.Post(
                _ =>
                {
                    Application.ExitThread();
                },
                null
            );
        }

        if (_thread.IsAlive)
            _thread.Join(TimeSpan.FromSeconds(5));

        _syncCtx?.Dispose();
        _ready.Dispose();
    }
}
