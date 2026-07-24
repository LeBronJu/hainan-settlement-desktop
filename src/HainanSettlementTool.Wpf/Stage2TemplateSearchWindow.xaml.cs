using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace HainanSettlementTool.Wpf
{
    public partial class Stage2TemplateSearchWindow : Window
    {
        private readonly Stage2TemplateSearchViewModel _viewModel;

        internal Stage2TemplateSearchWindow(
            IEnumerable<Stage2PreflightTemplateOptionViewModel> options,
            string selectedTemplatePath = null)
        {
            _viewModel = new Stage2TemplateSearchViewModel(options);
            InitializeComponent();
            DataContext = _viewModel;

            if (!string.IsNullOrWhiteSpace(selectedTemplatePath))
            {
                _viewModel.SelectedOption = _viewModel.Results.FirstOrDefault(option =>
                    string.Equals(option.Path, selectedTemplatePath, StringComparison.OrdinalIgnoreCase));
            }

            Loaded += (sender, args) =>
            {
                SearchBox.Focus();
                SearchBox.SelectAll();
                if (_viewModel.SelectedOption != null)
                {
                    ResultsList.ScrollIntoView(_viewModel.SelectedOption);
                }
            };
        }

        internal Stage2PreflightTemplateOptionViewModel SelectedOption
        {
            get { return _viewModel.SelectedOption; }
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            ConfirmSelection();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ResultsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ConfirmSelection();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Down && SearchBox.IsKeyboardFocusWithin)
            {
                if (_viewModel.SelectedOption == null)
                {
                    _viewModel.SelectedOption = _viewModel.Results.FirstOrDefault();
                }

                ResultsList.Focus();
                if (_viewModel.SelectedOption != null)
                {
                    ResultsList.ScrollIntoView(_viewModel.SelectedOption);
                }

                e.Handled = true;
                return;
            }

            if (e.Key == Key.Enter
                && ShouldConfirmFromEnter(
                    SearchBox.IsKeyboardFocusWithin,
                    ResultsList.IsKeyboardFocusWithin))
            {
                if (_viewModel.SelectedOption == null)
                {
                    _viewModel.SelectedOption = _viewModel.Results.FirstOrDefault();
                }

                ConfirmSelection();
                e.Handled = true;
            }
        }

        internal static bool ShouldConfirmFromEnter(
            bool searchBoxHasFocus,
            bool resultsListHasFocus)
        {
            return searchBoxHasFocus || resultsListHasFocus;
        }

        private void ConfirmSelection()
        {
            if (_viewModel.SelectedOption == null)
            {
                return;
            }

            DialogResult = true;
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }
}
