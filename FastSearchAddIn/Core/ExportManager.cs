using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastSearchAddIn.Core
{
    /// <summary>
    /// Exports selected emails (and their attachments) to the Windows Desktop,
    /// organised by a user-chosen sort criterion.
    /// </summary>
    public sealed class ExportManager
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        /// <summary>Root folder created under the Desktop.</summary>
        private const string ExportRootName = "Outlook_Exports";

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Exports the given <paramref name="results"/> to the Desktop asynchronously.
        /// </summary>
        /// <param name="results">Search results to export.</param>
        /// <param name="criterion">How to organise subfolders.</param>
        /// <param name="outlookApp">Live Outlook application instance.</param>
        /// <param name="progress">Optional progress callback (0–100).</param>
        /// <param name="token">Cancellation token.</param>
        /// <returns>Path to the export root folder.</returns>
        public Task<ExportSummary> ExportAsync(
            IReadOnlyList<SearchResult> results,
            ExportSortCriterion criterion,
            Outlook.Application outlookApp,
            IProgress<int> progress = null,
            CancellationToken token = default)
        {
            return Task.Run(() => Export(results, criterion, outlookApp, progress, token), token);
        }

        // -----------------------------------------------------------------------
        // Private – core export logic
        // -----------------------------------------------------------------------

        private ExportSummary Export(
            IReadOnlyList<SearchResult> results,
            ExportSortCriterion criterion,
            Outlook.Application outlookApp,
            IProgress<int> progress,
            CancellationToken token)
        {
            if (results == null || results.Count == 0)
                return new ExportSummary(0, 0, null, "No items selected.");

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string exportRoot  = Path.Combine(desktopPath, ExportRootName);
            Directory.CreateDirectory(exportRoot);

            int exported     = 0;
            int attachments  = 0;
            var errors       = new List<string>();

            for (int i = 0; i < results.Count; i++)
            {
                token.ThrowIfCancellationRequested();

                var result = results[i];
                Outlook.MailItem mail = null;
                try
                {
                    // Locate the MailItem in Outlook by its EntryID.
                    mail = (Outlook.MailItem)outlookApp.Session.GetItemFromID(
                        result.EntryId, result.StoreId);

                    // Determine the destination subfolder.
                    string subfolder = ResolveSubfolder(mail, criterion);
                    string destDir   = Path.Combine(exportRoot, subfolder);
                    Directory.CreateDirectory(destDir);

                    // Save the email as .msg.
                    string msgPath = BuildUniquePath(destDir, SanitiseFileName(mail.Subject ?? "NoSubject"), ".msg");
                    mail.SaveAs(msgPath, Outlook.OlSaveAsType.olMSG);
                    exported++;

                    // Save each attachment alongside the .msg file.
                    Outlook.Attachments atts = null;
                    try
                    {
                        atts = mail.Attachments;
                        for (int a = 1; a <= atts.Count; a++)
                        {
                            Outlook.Attachment att = null;
                            try
                            {
                                att = atts[a];
                                // Skip hidden / embedded OLE objects (type 5 = olEmbeddeditem handled; 6 = OLE).
                                if (att.Type == Outlook.OlAttachmentType.olOLE) continue;

                                string attName = SanitiseFileName(att.FileName);
                                if (string.IsNullOrWhiteSpace(attName)) attName = $"attachment_{a}";
                                string attPath = BuildUniquePath(destDir, Path.GetFileNameWithoutExtension(attName),
                                    Path.GetExtension(attName));
                                att.SaveAsFile(attPath);
                                attachments++;
                            }
                            catch (System.Exception ex)
                            {
                                Log.Warn(ex, "Failed to save attachment {0} of '{1}'", a, result.Subject);
                            }
                            finally
                            {
                                ReleaseComObject(att);
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(atts);
                    }
                }
                catch (System.Exception ex) when (!(ex is OperationCanceledException))
                {
                    Log.Error(ex, "Failed to export mail '{0}'", result.Subject);
                    errors.Add($"'{result.Subject}': {ex.Message}");
                }
                finally
                {
                    ReleaseComObject(mail);
                    // Force GC every 50 items to release COM memory.
                    if (i % 50 == 49)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }

                // Report progress (0–100).
                progress?.Report((int)((i + 1) * 100.0 / results.Count));
            }

            string summary = errors.Count == 0
                ? $"Exported {exported} email(s) and {attachments} attachment(s)."
                : $"Exported {exported}/{results.Count} email(s). {errors.Count} error(s) – see log.";

            Log.Info(summary);
            return new ExportSummary(exported, attachments, exportRoot, summary);
        }

        // -----------------------------------------------------------------------
        // Subfolder resolution
        // -----------------------------------------------------------------------

        private static string ResolveSubfolder(Outlook.MailItem mail, ExportSortCriterion criterion)
        {
            switch (criterion)
            {
                case ExportSortCriterion.Sender:
                    return SanitiseFileName(mail.SenderEmailAddress ?? "Unknown_Sender");

                case ExportSortCriterion.Date:
                    // e.g. "2025\04 April"
                    var dt   = mail.ReceivedTime;
                    string m = dt.ToString("MM MMMM", System.Globalization.CultureInfo.InvariantCulture);
                    return Path.Combine(dt.Year.ToString(), m);

                case ExportSortCriterion.Domain:
                    string domain = ExtractDomain(mail.SenderEmailAddress);
                    return SanitiseFileName(domain);

                default:
                    return "Unsorted";
            }
        }

        private static string ExtractDomain(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return "UnknownDomain";
            int at = email.IndexOf('@');
            if (at < 0 || at == email.Length - 1) return "UnknownDomain";
            return email.Substring(at + 1).ToLowerInvariant();
        }

        // -----------------------------------------------------------------------
        // File system helpers
        // -----------------------------------------------------------------------

        private static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();

        /// <summary>Replaces characters that are illegal in Windows file names.</summary>
        private static string SanitiseFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "_";
            foreach (char c in InvalidChars)
                name = name.Replace(c, '_');
            // Trim trailing dots/spaces (Windows quirk).
            name = name.TrimEnd('.', ' ');
            // Truncate to 100 chars to avoid MAX_PATH issues.
            if (name.Length > 100) name = name.Substring(0, 100);
            return string.IsNullOrWhiteSpace(name) ? "_" : name;
        }

        /// <summary>
        /// Builds a path that does not already exist by appending (1), (2), … as needed.
        /// </summary>
        private static string BuildUniquePath(string dir, string baseName, string ext)
        {
            string candidate = Path.Combine(dir, baseName + ext);
            if (!File.Exists(candidate)) return candidate;

            for (int i = 1; i < 10000; i++)
            {
                candidate = Path.Combine(dir, $"{baseName} ({i}){ext}");
                if (!File.Exists(candidate)) return candidate;
            }
            // Fallback with timestamp.
            return Path.Combine(dir, $"{baseName}_{DateTime.Now:HHmmssfff}{ext}");
        }

        // -----------------------------------------------------------------------
        // COM helper
        // -----------------------------------------------------------------------
        private static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try { Marshal.ReleaseComObject(obj); } catch { }
            }
        }
    }

    // -----------------------------------------------------------------------
    // Supporting types
    // -----------------------------------------------------------------------

    /// <summary>How exported emails are organised into subfolders.</summary>
    public enum ExportSortCriterion
    {
        Sender,
        Date,
        Domain,
    }

    /// <summary>Result returned after an export operation completes.</summary>
    public class ExportSummary
    {
        public int    EmailsExported     { get; }
        public int    AttachmentsExported { get; }
        public string ExportRootPath     { get; }
        public string Message            { get; }

        public ExportSummary(int emails, int attachments, string path, string message)
        {
            EmailsExported      = emails;
            AttachmentsExported = attachments;
            ExportRootPath      = path;
            Message             = message;
        }
    }
}
