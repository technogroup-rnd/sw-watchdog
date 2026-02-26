using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwWatchdog.Session;

/// <summary>
/// Implements the isolation protocol between sessions:
/// CloseAllDocuments → verify → SetSearchFolders → SetWorkingDir.
/// All methods MUST be called from the STA thread.
/// </summary>
internal sealed class IsolationProtocol
{
    private readonly ILogger _logger;

    public IsolationProtocol(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Run the full isolation protocol. Returns true if clean, false if tainted.
    /// </summary>
    public bool Isolate(ISldWorks swApp, string workingDirectory)
    {
        // Step 1: Close all documents (force, discard unsaved)
        swApp.CloseAllDocuments(true);

        // Step 2: Verify all closed
        var count = swApp.GetDocumentCount();
        if (count != 0)
        {
            _logger.LogWarning(
                "GetDocumentCount()={Count} after CloseAllDocuments, force-closing individually",
                count
            );

            // Step 3: Force-close individual documents
            ForceCloseAll(swApp);

            // Step 4: Re-verify
            count = swApp.GetDocumentCount();
            if (count != 0)
            {
                _logger.LogError(
                    "GetDocumentCount()={Count} after individual close — SW is tainted",
                    count
                );
                return false;
            }
        }

        // Step 5: Set search folders for referenced documents
        swApp.SetSearchFolders((int)swSearchFolderTypes_e.swDocumentType, workingDirectory);

        // Step 6: Set current working directory
        swApp.SetCurrentWorkingDirectory(workingDirectory);

        _logger.LogDebug("Isolation complete, workingDir={Dir}", workingDirectory);
        return true;
    }

    /// <summary>
    /// Run cleanup isolation (after session dispose). No directory setting needed.
    /// Returns true if clean, false if tainted.
    /// </summary>
    public bool Cleanup(ISldWorks swApp)
    {
        swApp.CloseAllDocuments(true);

        var count = swApp.GetDocumentCount();
        if (count != 0)
        {
            _logger.LogWarning(
                "Cleanup: GetDocumentCount()={Count} after CloseAllDocuments",
                count
            );
            ForceCloseAll(swApp);

            count = swApp.GetDocumentCount();
            if (count != 0)
            {
                _logger.LogError("Cleanup: GetDocumentCount()={Count} — SW is tainted", count);
                return false;
            }
        }

        return true;
    }

    private void ForceCloseAll(ISldWorks swApp)
    {
        var doc = (IModelDoc2?)swApp.GetFirstDocument();
        while (doc is not null)
        {
            var next = (IModelDoc2?)doc.GetNext();
            var title = doc.GetTitle();
            try
            {
                swApp.CloseDoc(title);
                _logger.LogDebug("Force-closed document: {Title}", title);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to close document: {Title}", title);
            }
            doc = next;
        }
    }
}
