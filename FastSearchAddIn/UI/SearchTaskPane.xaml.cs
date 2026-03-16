using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FastSearchAddIn.Core;
using NLog;

namespace FastSearchAddIn.UI
{
    // -------------------------------------------------------------------------
    // ViewModel (simple, no MVVM framework dependency)
    // -------------------------------------------------------------------------

    /// <summary>
    /// Lightweight ViewModel for the search task pane.
    /// Implements INotifyPropertyChanged so WPF bindings update automatically.
    /// </summary>
    public class SearchViewModel : INotifyPropertyChanged
    {
        private string _queryText = string.Empty;
        private ObservableCollection<SearchResult> _results = new ObservableCollection<SearchResult>();

        public string QueryText
        {
            get => _queryText;
            set { _queryText = value; OnPropertyChanged(); }
        }

        public ObservableCollection<SearchResult> Results
        {
            get => _results;
            set { _results = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // -------------------------------------------------------------------------
    // Code-behind
    // -------------------------------------------------------------------------

    /// <summary>
    /// WPF UserControl that hosts the search sidebar inside Outlook's task pane.
    /// </summary>
    public partial class SearchTaskPane : UserControl
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly IndexManager   _indexManager;
        private readonly SearchProvider _searchProvider;
        private readonly ExportManager  _exportManager;

        private readonly SearchViewModel _vm = new SearchViewModel();

        // Debounce timer for search-as-you-type.
        private CancellationTokenSource _searchCts;
        private readonly int _debounceMs;

        // Keep a reference to the Outlook app so we can export.
        private Microsoft.Office.Interop.Outlook.Application _outlookApp;

        // -----------------------------------------------------------------------
        // Constructor
        // -----------------------------------------------------------------------

        public SearchTaskPane(
            IndexManager   indexManager,
            SearchProvider searchProvider,
            ExportManager  exportManager)
        {
            _indexManager   = indexManager   ?? throw new ArgumentNullException(nameof(indexManager));
            _searchProvider = searchProvider ?? throw new ArgumentNullException(nameof(searchProvider));
            _exportManager  = exportManager  ?? throw new ArgumentNullException(nameof(exportManager));
            _debounceMs     = AppSettings.Current.SearchDebounceMs;

            InitializeComponent();
            DataContext = _vm;

            // Subscribe to indexing progress so we can update the status label.
            _indexManager.IndexProgressChanged += OnIndexProgressChanged;

            // Retrieve the Outlook Application from the add-in globals.
            _outlookApp = Globals.ThisAddIn.Application;
        }

        // -----------------------------------------------------------------------
        // Search events
        // -----------------------------------------------------------------------

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            TriggerDebouncedSearch(TxtSearch.Text);
        }

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                TxtSearch.Text = string.Empty;
                ClearResults();
            }
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            TxtSearch.Text = string.Empty;
            ClearResults();
            TxtSearch.Focus();
        }

        private void BtnReindex_Click(object sender, RoutedEventArgs e)
        {
            SetStatus("Re-indexing… this may take a while.", isError: false);
            PbarIndexing.Visibility = Visibility.Visible;
            BtnReindex.IsEnabled = false;
            _indexManager.StartIndexingAsync(_outlookApp);
        }

        // -----------------------------------------------------------------------
        // Export button events
        // -----------------------------------------------------------------------

        private void BtnExportSender_Click(object sender, RoutedEventArgs e)
            => TriggerExport(ExportSortCriterion.Sender);

        private void BtnExportDate_Click(object sender, RoutedEventArgs e)
            => TriggerExport(ExportSortCriterion.Date);

        private void BtnExportDomain_Click(object sender, RoutedEventArgs e)
            => TriggerExport(ExportSortCriterion.Domain);

        // -----------------------------------------------------------------------
        // Core logic – search
        // -----------------------------------------------------------------------

        private void TriggerDebouncedSearch(string query)
        {
            // Cancel any in-flight search.
            _searchCts?.Cancel();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;

            if (string.IsNullOrWhiteSpace(query))
            {
                ClearResults();
                return;
            }

            // Run on background thread; update UI via Dispatcher.
            Task.Run(async () =>
            {
                try
                {
                    var results = await _searchProvider.SearchAsync(query, _debounceMs, token)
                                                       .ConfigureAwait(false);
                    if (token.IsCancellationRequested) return;

                    Dispatcher.Invoke(() =>
                    {
                        _vm.Results.Clear();
                        foreach (var r in results)
                            _vm.Results.Add(r);

                        TxtResultCount.Text = results.Count > 0
                            ? $"{results.Count} result(s)"
                            : "No results";

                        if (!_indexManager.IsIndexReady)
                            SetStatus("Index is being built – results may be incomplete.", isError: false);
                        else
                            ClearStatus();
                    });
                }
                catch (OperationCanceledException) { /* debounced away */ }
                catch (Exception ex)
                {
                    Log.Error(ex, "Search failed.");
                    Dispatcher.Invoke(() => SetStatus($"Search error: {ex.Message}", isError: true));
                }
            }, token);
        }

        // -----------------------------------------------------------------------
        // Core logic – export
        // -----------------------------------------------------------------------

        private async void TriggerExport(ExportSortCriterion criterion)
        {
            var selected = LstResults.SelectedItems.Cast<SearchResult>().ToList();
            if (selected.Count == 0)
            {
                SetExportStatus("Select one or more emails from the list first.", isError: true);
                return;
            }

            SetExportStatus($"Exporting {selected.Count} email(s)…", isError: false);
            PbarExport.Value      = 0;
            PbarExport.Visibility = Visibility.Visible;
            SetExportButtonsEnabled(false);

            try
            {
                var progress = new Progress<int>(pct =>
                    Dispatcher.Invoke(() => PbarExport.Value = pct));

                var summary = await _exportManager.ExportAsync(
                    selected,
                    criterion,
                    _outlookApp,
                    progress,
                    CancellationToken.None).ConfigureAwait(false);

                Dispatcher.Invoke(() =>
                {
                    PbarExport.Value = 100;
                    SetExportStatus(summary.Message, isError: false);

                    // Offer to open the export folder.
                    if (summary.ExportRootPath != null &&
                        MessageBox.Show($"{summary.Message}\n\nOpen export folder?",
                            "Export Complete",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        Process.Start("explorer.exe", summary.ExportRootPath);
                    }
                });
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Export failed.");
                Dispatcher.Invoke(() =>
                    SetExportStatus($"Export failed – check log for details. ({ex.Message})", isError: true));
            }
            finally
            {
                Dispatcher.Invoke(() =>
                {
                    PbarExport.Visibility = Visibility.Collapsed;
                    SetExportButtonsEnabled(true);
                });
            }
        }

        // -----------------------------------------------------------------------
        // Index progress handler
        // -----------------------------------------------------------------------

        private void OnIndexProgressChanged(object sender, IndexProgressEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                switch (e.State)
                {
                    case IndexProgressState.Started:
                    case IndexProgressState.InProgress:
                        PbarIndexing.Visibility = Visibility.Visible;
                        SetStatus(e.Message, isError: false);
                        break;

                    case IndexProgressState.Completed:
                        PbarIndexing.Visibility = Visibility.Collapsed;
                        SetStatus(e.Message, isError: false);
                        BtnReindex.IsEnabled = true;
                        // Re-run current query now that the index is complete.
                        if (!string.IsNullOrWhiteSpace(TxtSearch.Text))
                            TriggerDebouncedSearch(TxtSearch.Text);
                        break;

                    case IndexProgressState.Cancelled:
                        PbarIndexing.Visibility = Visibility.Collapsed;
                        SetStatus("Indexing cancelled.", isError: false);
                        BtnReindex.IsEnabled = true;
                        break;

                    case IndexProgressState.Error:
                        PbarIndexing.Visibility = Visibility.Collapsed;
                        SetStatus(e.Message, isError: true);
                        BtnReindex.IsEnabled = true;
                        break;
                }
            });
        }

        // -----------------------------------------------------------------------
        // UI helpers
        // -----------------------------------------------------------------------

        private void ClearResults()
        {
            _vm.Results.Clear();
            TxtResultCount.Text = string.Empty;
        }

        private void ClearStatus()
        {
            TxtStatus.Text       = string.Empty;
            TxtStatus.Foreground = (System.Windows.Media.Brush)FindResource("StatusNormal");
        }

        private void SetStatus(string message, bool isError)
        {
            TxtStatus.Text       = message;
            TxtStatus.Foreground = isError
                ? (System.Windows.Media.Brush)FindResource("StatusError")
                : (System.Windows.Media.Brush)FindResource("StatusNormal");
        }

        private void SetExportStatus(string message, bool isError)
        {
            TxtExportStatus.Text       = message;
            TxtExportStatus.Foreground = isError
                ? (System.Windows.Media.Brush)FindResource("StatusError")
                : (System.Windows.Media.Brush)FindResource("StatusNormal");
        }

        private void SetExportButtonsEnabled(bool enabled)
        {
            BtnExportSender.IsEnabled = enabled;
            BtnExportDate.IsEnabled   = enabled;
            BtnExportDomain.IsEnabled = enabled;
        }
    }
}
