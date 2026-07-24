using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Wpf.Tests
{
    [TestClass]
    public class Stage2TemplateCandidateBrowserTests
    {
        [TestMethod]
        public void ThreeCandidatesAreAllVisibleWithoutBatchNavigation()
        {
            var browser = Browser(3, 101);

            Assert.AreEqual(3, browser.VisibleOptions.Count);
            Assert.AreEqual(Visibility.Collapsed, browser.NavigationVisibility);
            StringAssert.Contains(browser.BatchText, "已全部显示");
        }

        [TestMethod]
        public void SevenCandidatesUseFiveThenTwoWithoutRepeats()
        {
            var browser = Browser(7, 202);
            var firstBatch = browser.VisibleOptions.Select(option => option.Path).ToList();

            Assert.AreEqual(5, firstBatch.Count);
            browser.MoveNextOrReshuffle();
            var secondBatch = browser.VisibleOptions.Select(option => option.Path).ToList();

            Assert.AreEqual(2, secondBatch.Count);
            Assert.AreEqual(7, firstBatch.Concat(secondBatch).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.AreEqual("重新打乱", browser.NextButtonText);
            Assert.IsTrue(browser.CanMovePrevious);

            browser.MovePrevious();

            CollectionAssert.AreEqual(firstBatch, browser.VisibleOptions.Select(option => option.Path).ToList());
        }

        [TestMethod]
        public void ThirteenCandidatesUseFiveFiveThreeAndReshuffleAfterLastBatch()
        {
            var browser = Browser(13, 303);
            var firstRound = CollectCurrentRound(browser);

            CollectionAssert.AreEqual(new[] { 5, 5, 3 }, firstRound.BatchSizes.ToArray());
            Assert.AreEqual(13, firstRound.Paths.Distinct(StringComparer.OrdinalIgnoreCase).Count());

            browser.MoveNextOrReshuffle();
            var secondRound = CollectCurrentRound(browser);

            Assert.AreEqual(13, secondRound.Paths.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.IsFalse(firstRound.Paths.SequenceEqual(
                secondRound.Paths,
                StringComparer.OrdinalIgnoreCase));
        }

        [TestMethod]
        public void SelectionSurvivesBatchChangesAndRejectsUnknownPath()
        {
            var browser = Browser(7, 404);
            var selected = browser.VisibleOptions[0];

            Assert.IsTrue(browser.SelectTemplate(selected.Path));
            browser.MoveNextOrReshuffle();

            Assert.AreEqual(selected.Path, browser.SelectedTemplatePath);
            Assert.AreSame(selected, browser.SelectedOption);
            Assert.IsTrue(selected.IsSelected);
            Assert.AreEqual(Visibility.Visible, browser.SelectedOptionVisibility);
            Assert.IsFalse(browser.SelectTemplate(@"C:\templates\unknown.xlsx"));
            Assert.AreEqual(selected.Path, browser.SelectedTemplatePath);
        }

        [TestMethod]
        public void SearchUsesWhitespaceTermsAsCaseInsensitiveAndAcrossAllFields()
        {
            var first = Option(
                @"C:\templates\owner-alpha\subject-one.xlsx",
                "Subject Alpha",
                "Owner One",
                "first-template.xlsx");
            var second = Option(
                @"C:\templates\owner-beta\subject-two.xlsx",
                "Subject Alpha",
                "Owner Two",
                "second-template.xlsx");
            var third = Option(
                @"C:\templates\owner-three\subject-three.xlsx",
                "Subject Gamma",
                "Owner Three",
                "alpha-file.xlsx");
            var search = new Stage2TemplateSearchViewModel(new[] { first, second, third });

            search.Query = "ALPHA two";

            Assert.AreEqual(1, search.Results.Count);
            Assert.AreSame(second, search.Results.Single());
            Assert.AreEqual("找到 1 个模板", search.ResultCountText);

            search.Query = "alpha file";

            Assert.AreEqual(1, search.Results.Count);
            Assert.AreSame(third, search.Results.Single());
        }

        [TestMethod]
        public void SearchEnterDoesNotOverrideDialogButtons()
        {
            Assert.IsTrue(Stage2TemplateSearchWindow.ShouldConfirmFromEnter(
                searchBoxHasFocus: true,
                resultsListHasFocus: false));
            Assert.IsTrue(Stage2TemplateSearchWindow.ShouldConfirmFromEnter(
                searchBoxHasFocus: false,
                resultsListHasFocus: true));
            Assert.IsFalse(Stage2TemplateSearchWindow.ShouldConfirmFromEnter(
                searchBoxHasFocus: false,
                resultsListHasFocus: false));
        }

        private static Stage2TemplateCandidateBrowserViewModel Browser(int count, int seed)
        {
            return new Stage2TemplateCandidateBrowserViewModel(
                Enumerable.Range(1, count)
                    .Select(index => Option(
                        @"C:\templates\owner-" + index + @"\subject-" + index + ".xlsx",
                        "Subject " + index,
                        "Owner " + index,
                        "subject-" + index + ".xlsx")),
                random: new Random(seed));
        }

        private static Stage2PreflightTemplateOptionViewModel Option(
            string path,
            string subject,
            string owner,
            string fileName)
        {
            return new Stage2PreflightTemplateOptionViewModel
            {
                Path = path,
                SubjectText = subject,
                OwnerText = owner,
                FileName = fileName,
                DisplayText = subject + " / " + owner
            };
        }

        private static RoundSnapshot CollectCurrentRound(
            Stage2TemplateCandidateBrowserViewModel browser)
        {
            var paths = new List<string>();
            var batchSizes = new List<int>();
            while (true)
            {
                batchSizes.Add(browser.VisibleOptions.Count);
                paths.AddRange(browser.VisibleOptions.Select(option => option.Path));
                if (browser.NextButtonText == "重新打乱")
                {
                    break;
                }

                browser.MoveNextOrReshuffle();
            }

            return new RoundSnapshot(paths, batchSizes);
        }

        private sealed class RoundSnapshot
        {
            public RoundSnapshot(IReadOnlyList<string> paths, IReadOnlyList<int> batchSizes)
            {
                Paths = paths;
                BatchSizes = batchSizes;
            }

            public IReadOnlyList<string> Paths { get; }
            public IReadOnlyList<int> BatchSizes { get; }
        }
    }
}
