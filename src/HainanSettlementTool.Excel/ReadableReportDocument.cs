using System.Collections.Generic;

namespace HainanSettlementTool.Excel
{
    internal enum ReadableReportStatus
    {
        Success,
        Review,
        Critical
    }

    internal enum ReadableReportNoticeTone
    {
        Information,
        Review,
        Critical
    }

    internal sealed class ReadableReportDocument
    {
        public string Title { get; set; }

        public string PeriodLabel { get; set; }

        public ReadableReportStatus Status { get; set; }

        public string StatusText { get; set; }

        public string StatusDetail { get; set; }

        public string Footer { get; set; }

        public List<ReadableReportMetric> Metrics { get; } = new List<ReadableReportMetric>();

        public List<ReadableReportNotice> Notices { get; } = new List<ReadableReportNotice>();

        public List<ReadableReportSection> Sections { get; } = new List<ReadableReportSection>();
    }

    internal sealed class ReadableReportMetric
    {
        public ReadableReportMetric(string label, string value)
        {
            Label = label;
            Value = value;
        }

        public string Label { get; }

        public string Value { get; }
    }

    internal sealed class ReadableReportNotice
    {
        public ReadableReportNotice(
            string title,
            string body,
            ReadableReportNoticeTone tone)
        {
            Title = title;
            Body = body;
            Tone = tone;
        }

        public string Title { get; }

        public string Body { get; }

        public ReadableReportNoticeTone Tone { get; }
    }

    internal sealed class ReadableReportSection
    {
        public ReadableReportSection(string title, params string[] headers)
        {
            Title = title;
            Headers = new List<string>(headers ?? new string[0]);
        }

        public string Title { get; }

        public List<string> Headers { get; }

        public List<ReadableReportRow> Rows { get; } = new List<ReadableReportRow>();

        public string EmptyMessage { get; set; }

        public bool IsCollapsed { get; set; }
    }

    internal sealed class ReadableReportRow
    {
        public ReadableReportRow(params string[] cells)
        {
            Cells = new List<string>(cells ?? new string[0]);
        }

        public List<string> Cells { get; }
    }
}
