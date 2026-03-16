using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Lucene.Net.Index;
using Lucene.Net.QueryParsers.Classic;
using Lucene.Net.Search;
using Lucene.Net.Util;
using NLog;

namespace FastSearchAddIn.Core
{
    /// <summary>
    /// Executes real-time Lucene queries against the index built by
    /// <see cref="IndexManager"/>.  Results are returned asynchronously
    /// so the UI thread is never blocked.
    /// </summary>
    public sealed class SearchProvider : IDisposable
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        private static readonly LuceneVersion LuceneVer = LuceneVersion.LUCENE_48;

        private readonly IndexManager _indexManager;

        // Fields searched by the MultiFieldQueryParser.
        private static readonly string[] SearchFields = new[]
        {
            IndexManager.FieldSubject,
            IndexManager.FieldBody,
            IndexManager.FieldSender,
            IndexManager.FieldSenderEmail,
        };

        // Debounce support – only one search is scheduled at a time.
        private CancellationTokenSource _debounceCts;

        public SearchProvider(IndexManager indexManager)
        {
            _indexManager = indexManager ?? throw new ArgumentNullException(nameof(indexManager));
        }

        // -----------------------------------------------------------------------
        // Public API
        // -----------------------------------------------------------------------

        /// <summary>
        /// Schedules a debounced search.  Cancels any pending search, waits
        /// <paramref name="debounceMs"/> ms, then executes the query.
        /// Returns an empty list if the index is not yet ready.
        /// </summary>
        public async Task<IReadOnlyList<SearchResult>> SearchAsync(
            string queryText,
            int debounceMs = 300,
            CancellationToken externalToken = default)
        {
            // Cancel the previous search and dispose its token source to avoid a leak.
            var old = _debounceCts;
            old?.Cancel();
            _debounceCts = CancellationTokenSource.CreateLinkedTokenSource(externalToken);
            var token = _debounceCts.Token;
            old?.Dispose();

            try
            {
                // Debounce delay.
                await Task.Delay(debounceMs, token).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();

                // Execute on a pool thread.
                return await Task.Run(() => ExecuteSearch(queryText, token), token)
                                 .ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return Array.Empty<SearchResult>();
            }
        }

        /// <summary>
        /// Synchronous search, useful for callers that already manage threading.
        /// </summary>
        public IReadOnlyList<SearchResult> Search(string queryText)
        {
            return ExecuteSearch(queryText, CancellationToken.None);
        }

        // -----------------------------------------------------------------------
        // Private – query execution
        // -----------------------------------------------------------------------

        private IReadOnlyList<SearchResult> ExecuteSearch(string queryText, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(queryText))
                return Array.Empty<SearchResult>();

            try
            {
                using (var reader = DirectoryReader.Open(_indexManager.Directory))
                {
                    token.ThrowIfCancellationRequested();

                    var searcher = new IndexSearcher(reader);
                    var parser   = new MultiFieldQueryParser(LuceneVer, SearchFields, _indexManager.Analyzer)
                    {
                        DefaultOperator = QueryParserBase.Operator.OR,
                        AllowLeadingWildcard = true,
                    };

                    Query query;
                    try
                    {
                        // Support trailing wildcard for "search as you type".
                        string escaped = AppendWildcard(queryText);
                        query = parser.Parse(escaped);
                    }
                    catch (ParseException)
                    {
                        // If parsing fails (e.g. user typed an unclosed quote), fall back to
                        // a phrase query on the raw text.
                        query = parser.Parse(QueryParserBase.Escape(queryText));
                    }

                    int maxHits = AppSettings.Current.MaxResults;
                    TopDocs topDocs = searcher.Search(query, maxHits);

                    var results = new List<SearchResult>(topDocs.ScoreDocs.Length);
                    foreach (var scoreDoc in topDocs.ScoreDocs)
                    {
                        token.ThrowIfCancellationRequested();
                        var doc = searcher.Doc(scoreDoc.Doc);
                        // Document.Get() returns null for missing/unstored fields – default to empty string.
                        results.Add(new SearchResult
                        {
                            EntryId      = doc.Get(IndexManager.FieldEntryId)      ?? string.Empty,
                            StoreId      = doc.Get(IndexManager.FieldStoreId)      ?? string.Empty,
                            Subject      = doc.Get(IndexManager.FieldSubject)      ?? string.Empty,
                            SenderName   = doc.Get(IndexManager.FieldSender)       ?? string.Empty,
                            SenderEmail  = doc.Get(IndexManager.FieldSenderEmail)  ?? string.Empty,
                            ReceivedTime = ParseReceivedTime(doc.Get(IndexManager.FieldReceived)),
                            Score        = scoreDoc.Score,
                        });
                    }

                    return results;
                }
            }
            catch (System.IO.IOException ex)
            {
                // Index may not exist yet.
                Log.Debug(ex, "Index not available during search.");
                return Array.Empty<SearchResult>();
            }
            catch (System.Exception ex) when (!(ex is OperationCanceledException))
            {
                Log.Error(ex, "Search failed for query: {0}", queryText);
                return Array.Empty<SearchResult>();
            }
        }

        // -----------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------

        /// <summary>
        /// Appends a wildcard '*' to the last token so partial words are matched
        /// while the user is still typing.
        /// </summary>
        private static string AppendWildcard(string q)
        {
            q = q.Trim();
            if (q.Length == 0) return q;
            // Don't double-up wildcards or append after quotes.
            char last = q[q.Length - 1];
            if (last == '*' || last == '?' || last == '"') return q;
            return q + "*";
        }

        private static DateTime ParseReceivedTime(string s)
        {
            if (string.IsNullOrEmpty(s)) return DateTime.MinValue;
            if (DateTime.TryParseExact(s, "yyyyMMddHHmmss",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var dt))
                return dt;
            return DateTime.MinValue;
        }

        public void Dispose()
        {
            _debounceCts?.Cancel();
            _debounceCts?.Dispose();
        }
    }

    // -----------------------------------------------------------------------
    // DTO
    // -----------------------------------------------------------------------

    /// <summary>A single search hit returned to the UI.</summary>
    public class SearchResult
    {
        public string   EntryId      { get; set; }
        public string   StoreId      { get; set; }
        public string   Subject      { get; set; }
        public string   SenderName   { get; set; }
        public string   SenderEmail  { get; set; }
        public DateTime ReceivedTime { get; set; }
        public float    Score        { get; set; }

        /// <summary>Display string for the list box.</summary>
        public string DisplayLine =>
            $"{ReceivedTime:yyyy-MM-dd}  {SenderName}  –  {Subject}";
    }
}
