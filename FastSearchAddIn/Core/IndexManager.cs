using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Lucene.Net.Analysis.Standard;
using Lucene.Net.Documents;
using Lucene.Net.Index;
using Lucene.Net.Store;
using Lucene.Net.Util;
using NLog;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FastSearchAddIn.Core
{
    /// <summary>
    /// Builds and maintains a Lucene.NET index of the user's Outlook mailbox.
    /// Indexing runs on a background thread to keep Outlook responsive.
    /// </summary>
    public sealed class IndexManager : IDisposable
    {
        // -----------------------------------------------------------------------
        // Constants – Lucene field names
        // -----------------------------------------------------------------------
        public const string FieldEntryId  = "entryId";
        public const string FieldSubject  = "subject";
        public const string FieldBody     = "body";
        public const string FieldSender   = "sender";
        public const string FieldSenderEmail = "senderEmail";
        public const string FieldReceived = "received";   // stored as yyyyMMddHHmmss
        public const string FieldStoreId  = "storeId";

        // -----------------------------------------------------------------------
        // Fields
        // -----------------------------------------------------------------------
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        private static readonly LuceneVersion LuceneVer = LuceneVersion.LUCENE_48;

        private readonly FSDirectory _directory;
        private readonly StandardAnalyzer _analyzer;
        private IndexWriter _writer;
        private readonly object _writerLock = new object();

        private CancellationTokenSource _indexCts;
        private Task _indexTask;

        // Progress reporting – subscribers can listen to update the UI.
        public event EventHandler<IndexProgressEventArgs> IndexProgressChanged;

        // -----------------------------------------------------------------------
        // Constructor
        // -----------------------------------------------------------------------
        public IndexManager()
        {
            AppSettings.EnsureDirectories();
            _directory = FSDirectory.Open(AppSettings.IndexFolder);
            _analyzer  = new StandardAnalyzer(LuceneVer);
            OpenWriter();
        }

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>Returns true once the index has at least one document.</summary>
        public bool IsIndexReady
        {
            get
            {
                try
                {
                    using (var reader = DirectoryReader.Open(_directory))
                        return reader.NumDocs > 0;
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// Begins asynchronous full indexing of the Outlook mailbox.
        /// Safe to call multiple times – cancels any prior run.
        /// </summary>
        public void StartIndexingAsync(Outlook.Application outlookApp)
        {
            // Cancel any running indexing task.
            _indexCts?.Cancel();
            _indexCts = new CancellationTokenSource();
            var token = _indexCts.Token;

            _indexTask = Task.Run(() => RunIndexing(outlookApp, token), token);
        }

        /// <summary>
        /// Exposes the Lucene directory for read-only use by SearchProvider.
        /// </summary>
        public FSDirectory Directory => _directory;

        /// <summary>
        /// Exposes the analyzer for use by SearchProvider.
        /// </summary>
        public StandardAnalyzer Analyzer => _analyzer;

        // -----------------------------------------------------------------------
        // Private – indexing logic
        // -----------------------------------------------------------------------

        private void OpenWriter()
        {
            lock (_writerLock)
            {
                var config = new IndexWriterConfig(LuceneVer, _analyzer)
                {
                    OpenMode = OpenMode.CREATE_OR_APPEND
                };
                _writer = new IndexWriter(_directory, config);
            }
        }

        private void RunIndexing(Outlook.Application outlookApp, CancellationToken token)
        {
            Log.Info("Background indexing started.");
            RaiseProgress(IndexProgressState.Started, 0, 0, "Indexing started…");

            try
            {
                var stores = outlookApp.Session.Stores;
                int totalIndexed = 0;

                for (int s = 1; s <= stores.Count; s++)
                {
                    if (token.IsCancellationRequested) break;

                    Outlook.Store store = null;
                    Outlook.MAPIFolder rootFolder = null;
                    try
                    {
                        store = stores[s];
                        rootFolder = store.GetRootFolder();
                        IndexFolder(rootFolder, store.StoreID, ref totalIndexed, token);
                    }
                    catch (System.Exception ex)
                    {
                        Log.Warn(ex, "Skipping store {0}", s);
                    }
                    finally
                    {
                        ReleaseComObject(rootFolder);
                        ReleaseComObject(store);
                    }
                }

                // Commit all pending documents.
                lock (_writerLock) { _writer.Commit(); }

                RaiseProgress(IndexProgressState.Completed, totalIndexed, totalIndexed,
                    $"Indexing complete. {totalIndexed:N0} emails indexed.");
                Log.Info("Background indexing completed. {0} emails indexed.", totalIndexed);
            }
            catch (OperationCanceledException)
            {
                Log.Info("Indexing cancelled.");
                RaiseProgress(IndexProgressState.Cancelled, 0, 0, "Indexing cancelled.");
            }
            catch (System.Exception ex)
            {
                Log.Error(ex, "Indexing failed.");
                RaiseProgress(IndexProgressState.Error, 0, 0, $"Indexing error: {ex.Message}");
            }
            finally
            {
                ReleaseComObject(outlookApp.Session.Stores);
            }
        }

        private void IndexFolder(
            Outlook.MAPIFolder folder,
            string storeId,
            ref int totalIndexed,
            CancellationToken token)
        {
            if (token.IsCancellationRequested) return;

            // Skip non-mail folders (calendar, contacts, etc.).
            if (folder.DefaultItemType != Outlook.OlItemType.olMailItem)
                goto ProcessSubfolders;

            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                int count = items.Count;

                for (int i = 1; i <= count; i++)
                {
                    if (token.IsCancellationRequested) break;

                    object item = null;
                    try
                    {
                        item = items[i];
                        if (item is Outlook.MailItem mail)
                        {
                            IndexMailItem(mail, storeId);
                            totalIndexed++;

                            // Commit in batches to manage memory.
                            if (totalIndexed % 500 == 0)
                            {
                                lock (_writerLock) { _writer.Commit(); }
                                RaiseProgress(IndexProgressState.InProgress, totalIndexed, -1,
                                    $"Indexed {totalIndexed:N0} emails…");
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Log.Warn(ex, "Failed to index item {0} in folder {1}", i, folder.Name);
                    }
                    finally
                    {
                        ReleaseComObject(item);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Log.Warn(ex, "Failed to access items in folder {0}", folder.Name);
            }
            finally
            {
                ReleaseComObject(items);
            }

            ProcessSubfolders:
            // Recurse into subfolders.
            Outlook.Folders subfolders = null;
            try
            {
                subfolders = folder.Folders;
                for (int f = 1; f <= subfolders.Count; f++)
                {
                    if (token.IsCancellationRequested) break;
                    Outlook.MAPIFolder sub = null;
                    try
                    {
                        sub = subfolders[f];
                        IndexFolder(sub, storeId, ref totalIndexed, token);
                    }
                    finally
                    {
                        ReleaseComObject(sub);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Log.Warn(ex, "Failed to enumerate subfolders of {0}", folder.Name);
            }
            finally
            {
                ReleaseComObject(subfolders);
            }
        }

        private void IndexMailItem(Outlook.MailItem mail, string storeId)
        {
            string entryId       = mail.EntryID ?? string.Empty;
            string subject       = mail.Subject ?? string.Empty;
            string senderName    = mail.SenderName ?? string.Empty;
            string senderEmail   = mail.SenderEmailAddress ?? string.Empty;
            string body          = AppSettings.Current.IndexBody ? (mail.Body ?? string.Empty) : string.Empty;
            string receivedStr   = mail.ReceivedTime.ToString("yyyyMMddHHmmss");

            var doc = new Document
            {
                // Stored but not analysed – used to locate the email in Outlook.
                new StringField(FieldEntryId,     entryId,     Field.Store.YES),
                new StringField(FieldStoreId,     storeId,     Field.Store.YES),
                new StringField(FieldReceived,    receivedStr, Field.Store.YES),
                new StoredField(FieldSenderEmail, senderEmail),

                // Stored AND analysed for full-text search.
                new TextField(FieldSubject, subject,     Field.Store.YES),
                new TextField(FieldSender,  senderName,  Field.Store.YES),
            };

            if (AppSettings.Current.IndexBody)
                doc.Add(new TextField(FieldBody, body, Field.Store.NO));   // body not stored (large)

            lock (_writerLock)
            {
                // Update (delete-then-add) to avoid duplicates on re-index.
                _writer.UpdateDocument(new Term(FieldEntryId, entryId), doc);
            }
        }

        // -----------------------------------------------------------------------
        // Progress reporting
        // -----------------------------------------------------------------------
        private void RaiseProgress(IndexProgressState state, int indexed, int total, string message)
        {
            IndexProgressChanged?.Invoke(this, new IndexProgressEventArgs(state, indexed, total, message));
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

        // -----------------------------------------------------------------------
        // IDisposable
        // -----------------------------------------------------------------------
        public void Dispose()
        {
            _indexCts?.Cancel();
            try { _indexTask?.Wait(5000); } catch { }

            lock (_writerLock)
            {
                try { _writer?.Dispose(); } catch { }
            }
            try { _analyzer?.Dispose(); } catch { }
            try { _directory?.Dispose(); } catch { }
        }
    }

    // -----------------------------------------------------------------------
    // Supporting types
    // -----------------------------------------------------------------------

    public enum IndexProgressState { Started, InProgress, Completed, Cancelled, Error }

    public class IndexProgressEventArgs : EventArgs
    {
        public IndexProgressState State   { get; }
        public int                Indexed { get; }
        public int                Total   { get; }
        public string             Message { get; }

        public IndexProgressEventArgs(IndexProgressState state, int indexed, int total, string message)
        {
            State   = state;
            Indexed = indexed;
            Total   = total;
            Message = message;
        }
    }
}
