using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class EmployeeRewardService
    {
        private readonly IEmployeeRewardExcelGateway _excel;

        public EmployeeRewardService(IEmployeeRewardExcelGateway excel)
        {
            if (excel == null)
            {
                throw new ArgumentNullException(nameof(excel));
            }

            _excel = excel;
        }

        public EmployeeRewardResult Run(EmployeeRewardOptions options, Action<string> log)
        {
            ValidateOptions(options);
            Directory.CreateDirectory(options.OutputDirectory);

            log?.Invoke("正在读取员工电量奖励台账。");
            var ledgerRows = _excel.ReadLedgerRows(options);
            var result = BuildResult(options, ledgerRows);

            log?.Invoke("正在生成员工电量奖励总表和个人确认表。");
            var output = _excel.GenerateWorkbooks(options, result);
            result.SummaryPath = output?.SummaryPath;
            result.ReportPath = output?.ReportPath;
            result.PersonalWorkbookPaths = output?.PersonalWorkbookPaths ?? new List<string>();
            return result;
        }

        private static EmployeeRewardResult BuildResult(EmployeeRewardOptions options, IList<EmployeeRewardLedgerRow> ledgerRows)
        {
            var months = Enumerable.Range(options.StartMonth, options.EndMonth - options.StartMonth + 1).ToList();
            var rows = (ledgerRows ?? new List<EmployeeRewardLedgerRow>())
                .Where(row => !IsEmptyHelperRow(row))
                .Where(row => HasSelectedPower(row, months))
                .ToList();

            var errors = ValidateLedgerRows(rows, months);
            if (errors.Count > 0)
            {
                throw new InvalidOperationException("员工电量奖励无法生成，台账存在严重错误：" + Environment.NewLine + string.Join(Environment.NewLine, errors));
            }

            var details = rows
                .OrderBy(row => row.SourceRow)
                .Select(row => CreateDetail(row, months))
                .ToList();

            var summaries = details
                .GroupBy(row => row.Owner)
                .Select(group => CreateSummary(group.Key, group.ToList(), months))
                .ToList();

            var monthTotals = months.ToDictionary(month => month, month => Math.Round(details.Sum(row => row.MonthPowers[month]), 4));
            var totalPower = Math.Round(details.Sum(row => row.TotalPower), 4);

            return new EmployeeRewardResult
            {
                Year = options.Year,
                Months = months,
                Details = details,
                EmployeeSummaries = summaries,
                MonthTotals = monthTotals,
                TotalCustomers = details.Count,
                TotalPower = totalPower,
                TotalReward = totalPower
            };
        }

        private static EmployeeRewardDetail CreateDetail(EmployeeRewardLedgerRow row, IList<int> months)
        {
            var monthPowers = months.ToDictionary(month => month, month => Math.Round(GetPower(row, month), 4));
            return new EmployeeRewardDetail
            {
                SourceRow = row.SourceRow,
                Sequence = row.Sequence,
                CustomerCode = TextUtil.S(row.CustomerCode),
                CustomerName = TextUtil.S(row.CustomerName),
                ContractStartMonth = TextUtil.S(row.ContractStartMonth),
                Developer = TextUtil.S(row.Developer),
                AgentType = TextUtil.S(row.AgentType),
                Owner = TextUtil.S(row.Owner),
                MonthPowers = monthPowers,
                TotalPower = Math.Round(monthPowers.Values.Sum(), 4)
            };
        }

        private static EmployeeRewardSummary CreateSummary(string owner, IList<EmployeeRewardDetail> details, IList<int> months)
        {
            var monthPowers = months.ToDictionary(month => month, month => Math.Round(details.Sum(row => row.MonthPowers[month]), 4));
            var totalPower = Math.Round(monthPowers.Values.Sum(), 4);
            return new EmployeeRewardSummary
            {
                Owner = owner,
                CustomerCount = details.Count,
                MonthPowers = monthPowers,
                TotalPower = totalPower,
                RewardAmount = totalPower
            };
        }

        private static List<string> ValidateLedgerRows(IList<EmployeeRewardLedgerRow> rows, IList<int> months)
        {
            var errors = new List<string>();
            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.Owner))
                {
                    errors.Add("第" + row.SourceRow + "行负责人为空。");
                }

                if (string.IsNullOrWhiteSpace(row.CustomerName) && HasSelectedPower(row, months))
                {
                    errors.Add("第" + row.SourceRow + "行企业名称为空但有电量。");
                }
            }

            var duplicateCodes = rows
                .Where(row => !string.IsNullOrWhiteSpace(row.CustomerCode))
                .GroupBy(row => TextUtil.CustomerKey(row.CustomerCode))
                .Where(group => group.Count() > 1)
                .ToList();
            foreach (var group in duplicateCodes)
            {
                errors.Add("客户编号重复：" + group.First().CustomerCode + "，涉及台账行：" + string.Join("、", group.Select(row => row.SourceRow)));
            }

            return errors;
        }

        private static bool IsEmptyHelperRow(EmployeeRewardLedgerRow row)
        {
            return row == null
                || (string.IsNullOrWhiteSpace(row.CustomerCode)
                    && string.IsNullOrWhiteSpace(row.CustomerName)
                    && string.IsNullOrWhiteSpace(row.Owner));
        }

        private static bool HasSelectedPower(EmployeeRewardLedgerRow row, IList<int> months)
        {
            return months.Any(month => Math.Abs(GetPower(row, month)) > 0.0000001d);
        }

        private static double GetPower(EmployeeRewardLedgerRow row, int month)
        {
            if (row == null || row.MonthPowers == null)
            {
                return 0d;
            }

            double value;
            return row.MonthPowers.TryGetValue(month, out value) ? value : 0d;
        }

        private static void ValidateOptions(EmployeeRewardOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Year <= 0)
            {
                throw new ArgumentException("请设置奖励年份。");
            }

            if (options.StartMonth < 1 || options.StartMonth > 12 || options.EndMonth < 1 || options.EndMonth > 12)
            {
                throw new ArgumentException("奖励月份必须在1到12月之间。");
            }

            if (options.StartMonth > options.EndMonth)
            {
                throw new ArgumentException("开始月份不能晚于结束月份。");
            }

            FileAccessGuard.RequireReadableWorkbook(options.LedgerPath, "最新售电结算台账");

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择输出文件夹。");
            }
        }
    }
}
