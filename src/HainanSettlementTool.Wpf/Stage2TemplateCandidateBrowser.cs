using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace HainanSettlementTool.Wpf
{
    internal sealed class Stage2TemplateCandidateBrowserViewModel : INotifyPropertyChanged
    {
        internal const int BatchSize = 5;

        private readonly List<Stage2PreflightTemplateOptionViewModel> _allOptions;
        private readonly Random _random;
        private List<Stage2PreflightTemplateOptionViewModel> _roundOptions;
        private List<Stage2PreflightTemplateOptionViewModel> _visibleOptions;
        private int _batchIndex;
        private string _selectedTemplatePath;

        public Stage2TemplateCandidateBrowserViewModel(
            IEnumerable<Stage2PreflightTemplateOptionViewModel> options,
            string selectedTemplatePath = null,
            Random random = null)
        {
            _allOptions = (options ?? Enumerable.Empty<Stage2PreflightTemplateOptionViewModel>())
                .Where(option => option != null && !string.IsNullOrWhiteSpace(option.Path))
                .GroupBy(option => option.Path, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();
            _random = random ?? new Random(unchecked(Environment.TickCount * 397 ^ Guid.NewGuid().GetHashCode()));
            _roundOptions = new List<Stage2PreflightTemplateOptionViewModel>();
            _visibleOptions = new List<Stage2PreflightTemplateOptionViewModel>();

            StartNewRound();
            SelectTemplate(selectedTemplatePath);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public IReadOnlyList<Stage2PreflightTemplateOptionViewModel> AllOptions
        {
            get { return _allOptions; }
        }

        public IReadOnlyList<Stage2PreflightTemplateOptionViewModel> VisibleOptions
        {
            get { return _visibleOptions; }
        }

        public string SelectedTemplatePath
        {
            get { return _selectedTemplatePath; }
        }

        public Stage2PreflightTemplateOptionViewModel SelectedOption
        {
            get
            {
                return _allOptions.FirstOrDefault(option =>
                    string.Equals(option.Path, _selectedTemplatePath, StringComparison.OrdinalIgnoreCase));
            }
        }

        public Visibility SelectedOptionVisibility
        {
            get { return SelectedOption == null ? Visibility.Collapsed : Visibility.Visible; }
        }

        public Visibility NavigationVisibility
        {
            get { return _allOptions.Count > BatchSize ? Visibility.Visible : Visibility.Collapsed; }
        }

        public bool CanMovePrevious
        {
            get { return _batchIndex > 0; }
        }

        public string BatchText
        {
            get
            {
                if (_allOptions.Count == 0)
                {
                    return "没有可选模板";
                }

                if (_allOptions.Count <= BatchSize)
                {
                    return "共 " + _allOptions.Count + " 个候选，已全部显示";
                }

                return "第 " + (_batchIndex + 1) + " / " + BatchCount
                    + " 批 · 本批 " + _visibleOptions.Count + " 个 · 共 " + _allOptions.Count + " 个";
            }
        }

        public string NextButtonText
        {
            get { return IsLastBatch ? "重新打乱" : "换一批"; }
        }

        private int BatchCount
        {
            get { return (_allOptions.Count + BatchSize - 1) / BatchSize; }
        }

        private bool IsLastBatch
        {
            get { return _batchIndex >= Math.Max(0, BatchCount - 1); }
        }

        public bool SelectTemplate(string path)
        {
            Stage2PreflightTemplateOptionViewModel selected = null;
            if (!string.IsNullOrWhiteSpace(path))
            {
                selected = _allOptions.FirstOrDefault(option =>
                    string.Equals(option.Path, path, StringComparison.OrdinalIgnoreCase));
                if (selected == null)
                {
                    return false;
                }
            }

            _selectedTemplatePath = selected == null ? null : selected.Path;
            foreach (var option in _allOptions)
            {
                option.IsSelected = ReferenceEquals(option, selected);
            }

            OnPropertyChanged(nameof(SelectedTemplatePath));
            OnPropertyChanged(nameof(SelectedOption));
            OnPropertyChanged(nameof(SelectedOptionVisibility));
            return true;
        }

        public void MovePrevious()
        {
            if (!CanMovePrevious)
            {
                return;
            }

            _batchIndex--;
            RefreshVisibleBatch();
        }

        public void MoveNextOrReshuffle()
        {
            if (_allOptions.Count <= BatchSize)
            {
                return;
            }

            if (IsLastBatch)
            {
                StartNewRound();
                return;
            }

            _batchIndex++;
            RefreshVisibleBatch();
        }

        public void Reshuffle()
        {
            StartNewRound();
        }

        private void StartNewRound()
        {
            var previousOrder = _roundOptions
                .Select(option => option.Path)
                .ToList();
            _roundOptions = _allOptions.ToList();
            for (var index = _roundOptions.Count - 1; index > 0; index--)
            {
                var target = _random.Next(index + 1);
                var item = _roundOptions[index];
                _roundOptions[index] = _roundOptions[target];
                _roundOptions[target] = item;
            }

            if (_roundOptions.Count > 1
                && previousOrder.Count == _roundOptions.Count
                && previousOrder.SequenceEqual(
                    _roundOptions.Select(option => option.Path),
                    StringComparer.OrdinalIgnoreCase))
            {
                var first = _roundOptions[0];
                _roundOptions.RemoveAt(0);
                _roundOptions.Add(first);
            }

            _batchIndex = 0;
            RefreshVisibleBatch();
        }

        private void RefreshVisibleBatch()
        {
            _visibleOptions = _roundOptions
                .Skip(_batchIndex * BatchSize)
                .Take(BatchSize)
                .ToList();
            OnPropertyChanged(nameof(VisibleOptions));
            OnPropertyChanged(nameof(CanMovePrevious));
            OnPropertyChanged(nameof(BatchText));
            OnPropertyChanged(nameof(NextButtonText));
            OnPropertyChanged(nameof(NavigationVisibility));
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal sealed class Stage2TemplateSearchViewModel : INotifyPropertyChanged
    {
        private readonly IReadOnlyList<Stage2PreflightTemplateOptionViewModel> _allOptions;
        private string _query;
        private IReadOnlyList<Stage2PreflightTemplateOptionViewModel> _results;
        private Stage2PreflightTemplateOptionViewModel _selectedOption;

        public Stage2TemplateSearchViewModel(
            IEnumerable<Stage2PreflightTemplateOptionViewModel> options)
        {
            _allOptions = (options ?? Enumerable.Empty<Stage2PreflightTemplateOptionViewModel>())
                .Where(option => option != null)
                .ToList();
            _query = string.Empty;
            RefreshResults();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public string Query
        {
            get { return _query; }
            set
            {
                var normalized = value ?? string.Empty;
                if (string.Equals(_query, normalized, StringComparison.Ordinal))
                {
                    return;
                }

                _query = normalized;
                OnPropertyChanged();
                RefreshResults();
            }
        }

        public IReadOnlyList<Stage2PreflightTemplateOptionViewModel> Results
        {
            get { return _results; }
        }

        public string ResultCountText
        {
            get { return "找到 " + _results.Count + " 个模板"; }
        }

        public Stage2PreflightTemplateOptionViewModel SelectedOption
        {
            get { return _selectedOption; }
            set
            {
                if (ReferenceEquals(_selectedOption, value))
                {
                    return;
                }

                _selectedOption = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanConfirm));
            }
        }

        public bool CanConfirm
        {
            get { return SelectedOption != null; }
        }

        private void RefreshResults()
        {
            var terms = (_query ?? string.Empty)
                .Split(
                    new[] { ' ', '\t', '\r', '\n', '\u3000' },
                    StringSplitOptions.RemoveEmptyEntries);
            _results = _allOptions
                .Where(option => terms.All(term => MatchesAnyField(option, term)))
                .OrderBy(option => option.SubjectText, StringComparer.CurrentCulture)
                .ThenBy(option => option.OwnerText, StringComparer.CurrentCulture)
                .ThenBy(option => option.FileName, StringComparer.CurrentCulture)
                .ToList();

            if (_selectedOption != null && !_results.Contains(_selectedOption))
            {
                _selectedOption = null;
                OnPropertyChanged(nameof(SelectedOption));
                OnPropertyChanged(nameof(CanConfirm));
            }

            OnPropertyChanged(nameof(Results));
            OnPropertyChanged(nameof(ResultCountText));
        }

        private static bool MatchesAnyField(
            Stage2PreflightTemplateOptionViewModel option,
            string term)
        {
            return Contains(option.SubjectText, term)
                || Contains(option.OwnerText, term)
                || Contains(option.FileName, term);
        }

        private static bool Contains(string value, string term)
        {
            return !string.IsNullOrEmpty(value)
                && value.IndexOf(term, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
