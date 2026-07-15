using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongStage2WorkbookWriter
    {
        public IList<GuangdongStage2WorkbookResult> Write(
            string runDirectory,
            IEnumerable<GuangdongStage2WorkbookPreparation> preparations)
        {
            var results = new List<GuangdongStage2WorkbookResult>();
            foreach (var preparation in preparations)
            {
                results.Add(WriteOne(runDirectory, preparation));
            }

            return results;
        }

        private static GuangdongStage2WorkbookResult WriteOne(
            string runDirectory,
            GuangdongStage2WorkbookPreparation preparation)
        {
            var plan = preparation.Plan;
            var result = NewResult(plan);
            if (!plan.CanProcess)
            {
                string preservationError;
                if (!TryPreserveSourceWorkbook(runDirectory, plan, result, out preservationError))
                {
                    result.Action = GuangdongStage2PreparationActions.Failed;
                    result.IssueKind = GuangdongStage2IssueKinds.SkippedWorkbookPreservationFailed;
                    result.Message = "未自动处理，且原文件保留失败。原原因：" + plan.Message
                        + "；复制失败：" + preservationError;
                }

                return result;
            }

            var outputPath = Path.Combine(
                runDirectory,
                GuangdongStage2ExcelUtil.OutputFolderName(plan.SettlementKind),
                plan.RelativePath);
            result.OutputPath = outputPath;

            try
            {
                GuangdongStage2ExcelUtil.CopyWorkbookShared(plan.SourcePath, outputPath);
                if (plan.Action == GuangdongStage2PreparationActions.AlreadyPrepared)
                {
                    return result;
                }

                using (var workbook = new XLWorkbook(outputPath))
                {
                    IXLWorksheet worksheet;
                    if (plan.Action == GuangdongStage2PreparationActions.CreateTargetMonth)
                    {
                        worksheet = workbook.Worksheet(plan.SourceSheetName).CopyTo(plan.TargetSheetName);
                    }
                    else
                    {
                        worksheet = workbook.Worksheet(plan.TargetSheetName);
                    }

                    ApplyPreparation(worksheet, preparation.Layout, plan, result);
                    GuangdongStage2ExcelUtil.SaveWorkbook(workbook);
                }

                return result;
            }
            catch (Exception ex)
            {
                string partialOutputWarning = null;
                try
                {
                    if (File.Exists(outputPath))
                    {
                        File.Delete(outputPath);
                    }
                }
                catch (Exception cleanupException)
                {
                    partialOutputWarning = "；不完整输出未能删除，请勿使用：" + outputPath
                        + "（" + cleanupException.Message + "）";
                }

                result.Action = GuangdongStage2PreparationActions.Failed;
                result.IssueKind = GuangdongStage2IssueKinds.GenerationFailed;
                result.Message = "生成失败：" + ex.Message + partialOutputWarning;
                result.OutputPath = null;
                string preservationError;
                if (!TryPreserveSourceWorkbook(runDirectory, plan, result, out preservationError))
                {
                    result.Message += "；原文件保留失败：" + preservationError;
                }

                return result;
            }
        }

        private static bool TryPreserveSourceWorkbook(
            string runDirectory,
            GuangdongStage2WorkbookPlan plan,
            GuangdongStage2WorkbookResult result,
            out string error)
        {
            error = null;
            var reviewCopyPath = Path.Combine(
                runDirectory,
                GuangdongStage2ExcelUtil.OutputFolderName(plan.SettlementKind),
                GuangdongStage2ExcelUtil.ReviewFolderName,
                plan.RelativePath);
            try
            {
                GuangdongStage2ExcelUtil.CopyWorkbookShared(plan.SourcePath, reviewCopyPath);
                result.ReviewCopyPath = reviewCopyPath;
                return true;
            }
            catch (Exception ex)
            {
                result.ReviewCopyPath = null;
                error = ex.Message;
                return false;
            }
        }

        private static GuangdongStage2WorkbookResult NewResult(GuangdongStage2WorkbookPlan plan)
        {
            return new GuangdongStage2WorkbookResult
            {
                SettlementKind = plan.SettlementKind,
                SourcePath = plan.SourcePath,
                RelativePath = plan.RelativePath,
                Action = plan.Action,
                IssueKind = plan.IssueKind,
                Message = plan.Message,
                DetailRowCount = plan.DetailRowCount
            };
        }

        private static void ApplyPreparation(
            IXLWorksheet worksheet,
            GuangdongStage2WorksheetLayout layout,
            GuangdongStage2WorkbookPlan plan,
            GuangdongStage2WorkbookResult result)
        {
            if (plan.PowerNeedsClearing)
            {
                worksheet.Range(layout.DetailStartRow, 3, layout.DetailEndRow, 6)
                    .Clear(XLClearOptions.Contents);
                result.PowerCleared = true;
            }

            if (plan.TotalPowerNeedsReset)
            {
                for (var column = 3; column <= 6; column++)
                {
                    var cell = worksheet.Cell(layout.TotalRow, column);
                    if (!cell.HasFormula && !cell.IsEmpty())
                    {
                        cell.Value = 0;
                    }
                }

                result.TotalPowerReset = true;
            }

            if (plan.PeriodNeedsUpdate)
            {
                worksheet.Cell(layout.PeriodCellAddress).Value = layout.TargetPeriodText;
                result.PeriodUpdated = true;
            }

            if (plan.SettlementDateNeedsUpdate)
            {
                worksheet.Cell(layout.SettlementDateCellAddress).Value = layout.TargetSettlementDateText;
                result.SettlementDateUpdated = true;
            }
        }
    }
}
