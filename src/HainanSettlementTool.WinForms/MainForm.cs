using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Ookii.Dialogs.WinForms;

namespace HainanSettlementTool.WinForms
{
    public sealed class MainForm : Form
    {
        private readonly ComboBox _month = new ComboBox();
        private readonly TextBox _baseLedger = new TextBox();
        private readonly TextBox _power = new TextBox();
        private readonly TextBox _rawDetail = new TextBox();
        private readonly TextBox _referenceLedger = new TextBox();
        private readonly TextBox _completedLedger = new TextBox();
        private readonly TextBox _proxyTemplateDir = new TextBox();
        private readonly TextBox _intermediaryTemplateDir = new TextBox();
        private readonly TextBox _summaryTemplate = new TextBox();
        private readonly TextBox _outputDir = new TextBox();
        private readonly CheckBox _copyReferenceExisting = new CheckBox();
        private readonly CheckBox _allowMissingOwner = new CheckBox();
        private readonly Button _runStage1 = new Button();
        private readonly Button _cleanPower = new Button();
        private readonly Button _runStage2 = new Button();
        private readonly TextBox _log = new TextBox();
        private readonly Label _status = new Label();

        private static readonly Color WindowBackground = Color.FromArgb(246, 247, 249);
        private static readonly Color PanelBackground = Color.White;
        private static readonly Color BorderColor = Color.FromArgb(222, 226, 230);
        private static readonly Color MutedText = Color.FromArgb(92, 99, 112);
        private static readonly Color MainText = Color.FromArgb(34, 40, 49);
        private static readonly Color Accent = Color.FromArgb(0, 118, 125);
        private static readonly Color AccentHover = Color.FromArgb(0, 96, 102);

        public MainForm()
        {
            Text = "海南售电结算自动化工具";
            ClientSize = new Size(1180, 820);
            MinimumSize = new Size(1080, 820);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = WindowBackground;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            AutoScaleMode = AutoScaleMode.Dpi;

            BuildLayout();
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(22, 18, 22, 18),
                BackColor = WindowBackground
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            root.Controls.Add(BuildHeader(), 0, 0);

            var scrollHost = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = WindowBackground,
                Margin = new Padding(0)
            };
            root.Controls.Add(scrollHost, 0, 1);

            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                BackColor = WindowBackground,
                Margin = new Padding(0)
            };
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute, 140));
            scrollHost.Controls.Add(content);

            content.Controls.Add(BuildBody(), 0, 0);
            content.Controls.Add(BuildLogPanel(), 0, 1);
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 12)
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var titleBlock = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 1,
                AutoSize = true
            };
            titleBlock.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            titleBlock.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "海南售电结算自动化工具",
                ForeColor = MainText,
                Font = new Font(Font.FontFamily, 18F, FontStyle.Bold),
                Margin = new Padding(0)
            }, 0, 0);

            _status.AutoSize = true;
            _status.Text = "就绪";
            _status.ForeColor = Accent;
            _status.BackColor = Color.FromArgb(230, 246, 247);
            _status.Font = new Font(Font.FontFamily, 9F, FontStyle.Bold);
            _status.Padding = new Padding(14, 7, 14, 7);
            _status.Margin = new Padding(12, 8, 0, 0);
            _status.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            header.Controls.Add(titleBlock, 0, 0);
            header.Controls.Add(_status, 1, 0);
            return header;
        }

        private Control BuildBody()
        {
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 16)
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));

            var inputPanel = BuildInputSection();
            inputPanel.Margin = new Padding(0, 0, 10, 0);

            var workflowPanel = CreatePanel();
            workflowPanel.Margin = new Padding(10, 0, 0, 0);
            workflowPanel.Controls.Add(BuildWorkflowSection());

            body.Controls.Add(inputPanel, 0, 0);
            body.Controls.Add(workflowPanel, 1, 0);
            return body;
        }

        private Control BuildInputSection()
        {
            var section = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                BackColor = WindowBackground,
                Margin = new Padding(0)
            };
            section.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            section.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var commonForm = CreateFormTable("公共设置");
            _month.DropDownStyle = ComboBoxStyle.DropDownList;
            for (var month = 2; month <= 12; month++)
            {
                _month.Items.Add("2026年" + month + "月");
            }
            _month.SelectedIndex = -1;
            _month.Height = 30;
            AddControlRow(commonForm, "结算月份", _month, null);
            AddFolderRow(commonForm, "结果输出文件夹", _outputDir);
            section.Controls.Add(WrapCard(commonForm, new Padding(0, 0, 0, 12)), 0, 0);

            var stages = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                BackColor = WindowBackground,
                Margin = new Padding(0)
            };
            stages.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            stages.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            var stage1Form = CreateStackFormTable("阶段一：写入电量到台账");
            AddStackPathRow(stage1Form, "基础台账（必填）", _baseLedger, "Excel 文件|*.xlsx");
            AddStackPathRow(stage1Form, "电量处理表", _power, "Excel 文件|*.xlsx");
            AddStackPathRow(stage1Form, "原始零售侧明细", _rawDetail, "Excel/CSV|*.xlsx;*.xls;*.csv|Excel 文件|*.xlsx;*.xls|CSV 文件|*.csv|所有文件|*.*");
            AddStackPathRow(stage1Form, "参考台账（可选）", _referenceLedger, "Excel 文件|*.xlsx");

            _copyReferenceExisting.Text = "用参考台账覆盖已有客户 B:AB 列资料（谨慎）";
            AddStackCheckRow(stage1Form, _copyReferenceExisting);

            _runStage1.Text = "阶段一 写入电量到台账";
            _runStage1.Width = 205;
            _runStage1.Height = 38;
            StylePrimaryButton(_runStage1);
            _runStage1.Click += async (sender, args) => await RunStage1Async();
            AddStackActionRow(stage1Form, _runStage1);

            _cleanPower.Text = "只清洗电量数据";
            _cleanPower.Width = 160;
            _cleanPower.Height = 34;
            StyleSecondaryButton(_cleanPower);
            _cleanPower.Click += async (sender, args) => await RunCleanPowerAsync();
            AddStackActionRow(stage1Form, _cleanPower);

            var stage1Card = WrapCard(stage1Form, new Padding(0, 0, 6, 0));
            stages.Controls.Add(stage1Card, 0, 0);

            var stage2Form = CreateStackFormTable("阶段二：生成分表和汇总表");
            AddStackPathRow(stage2Form, "人工整理后的台账", _completedLedger, "Excel 文件|*.xlsx");
            AddStackFolderRow(stage2Form, "上月代理分表文件夹", _proxyTemplateDir);
            AddStackFolderRow(stage2Form, "上月居间分表文件夹", _intermediaryTemplateDir);
            AddStackPathRow(stage2Form, "上月/修正版汇总表", _summaryTemplate, "Excel 文件|*.xlsx");

            _allowMissingOwner.Text = "允许负责人缺失继续生成";
            AddStackCheckRow(stage2Form, _allowMissingOwner);

            _runStage2.Text = "阶段二 生成分表和汇总表";
            _runStage2.Width = 220;
            _runStage2.Height = 38;
            StylePrimaryButton(_runStage2);
            _runStage2.Click += async (sender, args) => await RunStage2Async();
            AddStackActionRow(stage2Form, _runStage2);

            var stage2Card = WrapCard(stage2Form, new Padding(6, 0, 0, 0));
            stages.Controls.Add(stage2Card, 1, 0);
            section.Controls.Add(stages, 0, 1);

            return section;
        }

        private Control WrapCard(Control content, Padding margin)
        {
            var panel = CreatePanel();
            panel.Margin = margin;
            panel.Controls.Add(content);
            return panel;
        }

        private TableLayoutPanel CreateFormTable(string title)
        {
            var form = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                RowCount = 1,
                AutoSize = true,
                Padding = new Padding(18, 14, 18, 16),
                BackColor = PanelBackground,
                Margin = new Padding(0)
            };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            form.Controls.Add(SectionTitle(title), 0, 0);
            form.SetColumnSpan(form.GetControlFromPosition(0, 0), 3);
            return form;
        }

        private TableLayoutPanel CreateStackFormTable(string title)
        {
            var form = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 1,
                AutoSize = true,
                Padding = new Padding(16, 14, 16, 16),
                BackColor = PanelBackground,
                Margin = new Padding(0)
            };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(SectionTitle(title), 0, 0);
            return form;
        }

        private static void AddStackPathRow(TableLayoutPanel form, string label, TextBox textBox, string filter)
        {
            var button = CreateBrowseButton();
            button.Click += (sender, args) =>
            {
                using (var dialog = new OpenFileDialog { Filter = filter })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox.Text = dialog.FileName;
                    }
                }
            };
            AddStackPickerRow(form, label, textBox, button);
        }

        private static void AddStackFolderRow(TableLayoutPanel form, string label, TextBox textBox)
        {
            var button = CreateBrowseButton();
            button.Click += (sender, args) =>
            {
                using (var dialog = new VistaFolderBrowserDialog
                {
                    Description = label,
                    UseDescriptionForTitle = true,
                    ShowNewFolderButton = true
                })
                {
                    var current = textBox.Text.Trim();
                    if (Directory.Exists(current))
                    {
                        dialog.SelectedPath = current;
                    }

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox.Text = dialog.SelectedPath;
                    }
                }
            };
            AddStackPickerRow(form, label, textBox, button);
        }

        private static void AddStackPickerRow(TableLayoutPanel form, string label, TextBox textBox, Button button)
        {
            var labelRow = AddRow(form);
            form.Controls.Add(new Label
            {
                Text = label,
                AutoSize = true,
                ForeColor = MutedText,
                Margin = new Padding(0, 3, 0, 0)
            }, 0, labelRow);

            var picker = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                BackColor = PanelBackground,
                Margin = new Padding(0, 0, 0, 1)
            };
            picker.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            picker.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82));
            textBox.Dock = DockStyle.Fill;
            textBox.Margin = new Padding(0, 2, 6, 2);
            StyleInput(textBox);
            button.Margin = new Padding(0, 2, 0, 2);
            picker.Controls.Add(textBox, 0, 0);
            picker.Controls.Add(button, 1, 0);

            var pickerRow = AddRow(form);
            form.Controls.Add(picker, 0, pickerRow);
        }

        private static void AddStackCheckRow(TableLayoutPanel form, CheckBox checkBox)
        {
            checkBox.ForeColor = MainText;
            checkBox.AutoSize = true;
            checkBox.Margin = new Padding(0, 5, 0, 4);
            var row = AddRow(form);
            form.Controls.Add(checkBox, 0, row);
        }

        private static void AddStackActionRow(TableLayoutPanel form, Button button)
        {
            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0, 2, 0, 0)
            };
            actions.Controls.Add(button);
            var row = AddRow(form);
            form.Controls.Add(actions, 0, row);
        }

        private Control BuildWorkflowSection()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                AutoSize = true,
                Padding = new Padding(20, 18, 20, 20),
                BackColor = PanelBackground
            };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            panel.Controls.Add(SectionTitle("使用提示"), 0, 0);
            panel.Controls.Add(WorkflowItem("1", "结果输出文件夹", "生成文件都会放在这里。"), 0, 1);
            panel.Controls.Add(WorkflowItem("2", "只跑阶段一", "选择台账和电量来源。"), 0, 2);
            panel.Controls.Add(WorkflowItem("3", "只跑阶段二", "选择整理台账和上月模板。"), 0, 3);
            panel.Controls.Add(WorkflowItem("4", "运行前关闭 Excel", "目标文件打开时会提示关闭。"), 0, 4);
            panel.Controls.Add(WorkflowItem("5", "生成后先检查", "确认汇总表和分表公式。"), 0, 5);

            return panel;
        }

        private Control BuildLogPanel()
        {
            var panel = CreatePanel();
            panel.Dock = DockStyle.Fill;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(20, 16, 20, 20),
                BackColor = PanelBackground
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.Controls.Add(layout);

            layout.Controls.Add(SectionTitle("运行日志"), 0, 0);

            _log.Multiline = true;
            _log.ScrollBars = ScrollBars.Vertical;
            _log.Dock = DockStyle.Fill;
            _log.ReadOnly = true;
            _log.BorderStyle = BorderStyle.None;
            _log.BackColor = Color.FromArgb(31, 35, 42);
            _log.ForeColor = Color.FromArgb(234, 238, 243);
            _log.Font = new Font("Consolas", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            _log.Margin = new Padding(0, 8, 0, 0);
            layout.Controls.Add(_log, 0, 1);

            return panel;
        }

        private static Panel CreatePanel()
        {
            return new ModernPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = PanelBackground,
                Padding = new Padding(1),
                Margin = new Padding(0)
            };
        }

        private sealed class ModernPanel : Panel
        {
            public ModernPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(BorderColor))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                }
            }
        }

        private Label SectionTitle(string text)
        {
            return new Label
            {
                AutoSize = true,
                Text = text,
                ForeColor = MainText,
                Font = new Font(Font.FontFamily, 11F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 10)
            };
        }

        private Control WorkflowItem(string number, string title, string detail)
        {
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Margin = new Padding(0, 0, 0, 15),
                BackColor = PanelBackground
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 34));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var badge = new Label
            {
                Text = number,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.White,
                BackColor = Accent,
                Width = 26,
                Height = 26,
                Font = new Font(Font.FontFamily, 9F, FontStyle.Bold),
                Margin = new Padding(0, 1, 8, 0)
            };

            var text = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = PanelBackground
            };
            text.Controls.Add(new Label
            {
                AutoSize = true,
                Text = title,
                ForeColor = MainText,
                Font = new Font(Font.FontFamily, 9.5F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 3)
            }, 0, 0);
            text.Controls.Add(new Label
            {
                AutoSize = true,
                Text = detail,
                ForeColor = MutedText,
                MaximumSize = new Size(260, 0),
                Margin = new Padding(0)
            }, 0, 1);

            row.Controls.Add(badge, 0, 0);
            row.Controls.Add(text, 1, 0);
            return row;
        }

        private static void AddControlRow(TableLayoutPanel form, string label, Control control, Button button)
        {
            var row = AddRow(form);
            form.Controls.Add(new Label
            {
                Text = label,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                ForeColor = MutedText,
                Margin = new Padding(0, 8, 14, 8)
            }, 0, row);
            control.Dock = DockStyle.Fill;
            control.Margin = new Padding(0, 5, 10, 5);
            StyleInput(control);
            form.Controls.Add(control, 1, row);
            if (button != null)
            {
                button.Margin = new Padding(0, 5, 0, 5);
                form.Controls.Add(button, 2, row);
            }
            else
            {
                form.Controls.Add(new Panel { Dock = DockStyle.Fill, Height = 1 }, 2, row);
            }
        }

        private static int AddRow(TableLayoutPanel form)
        {
            var row = form.RowCount;
            form.RowCount = row + 1;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return row;
        }

        private static void AddPathRow(TableLayoutPanel form, string label, TextBox textBox, string filter)
        {
            var button = CreateBrowseButton();
            button.Click += (sender, args) =>
            {
                using (var dialog = new OpenFileDialog { Filter = filter })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox.Text = dialog.FileName;
                    }
                }
            };
            AddControlRow(form, label, textBox, button);
        }

        private static void AddFolderRow(TableLayoutPanel form, string label, TextBox textBox)
        {
            var button = CreateBrowseButton();
            button.Click += (sender, args) =>
            {
                using (var dialog = new VistaFolderBrowserDialog
                {
                    Description = label,
                    UseDescriptionForTitle = true,
                    ShowNewFolderButton = true
                })
                {
                    var current = textBox.Text.Trim();
                    if (Directory.Exists(current))
                    {
                        dialog.SelectedPath = current;
                    }

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox.Text = dialog.SelectedPath;
                    }
                }
            };
            AddControlRow(form, label, textBox, button);
        }

        private static Button CreateBrowseButton()
        {
            var button = new Button
            {
                Text = "浏览",
                Dock = DockStyle.Fill,
                Height = 30
            };
            StyleSecondaryButton(button);
            return button;
        }

        private static void StyleInput(Control control)
        {
            var textBox = control as TextBox;
            if (textBox != null)
            {
                textBox.BorderStyle = BorderStyle.FixedSingle;
                textBox.BackColor = Color.White;
                textBox.ForeColor = MainText;
                textBox.Height = 30;
                return;
            }

            var comboBox = control as ComboBox;
            if (comboBox != null)
            {
                comboBox.FlatStyle = FlatStyle.Flat;
                comboBox.BackColor = Color.White;
                comboBox.ForeColor = MainText;
            }
        }

        private static void StylePrimaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Accent;
            button.ForeColor = Color.White;
            button.Font = new Font(button.Font.FontFamily, 9F, FontStyle.Bold);
            button.Cursor = Cursors.Hand;
            button.MouseEnter += (sender, args) => button.BackColor = AccentHover;
            button.MouseLeave += (sender, args) => button.BackColor = Accent;
        }

        private static void StyleSecondaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = BorderColor;
            button.FlatAppearance.BorderSize = 1;
            button.BackColor = Color.FromArgb(250, 251, 252);
            button.ForeColor = MainText;
            button.Cursor = Cursors.Hand;
            button.MouseEnter += (sender, args) => button.BackColor = Color.FromArgb(241, 244, 247);
            button.MouseLeave += (sender, args) => button.BackColor = Color.FromArgb(250, 251, 252);
        }

        private async Task RunStage1Async()
        {
            Stage1Options options;
            try
            {
                options = CreateOptions();
                if (!ConfirmRun("阶段一", options.Month, options.OutputDirectory))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
                return;
            }

            _runStage1.Enabled = false;
            _cleanPower.Enabled = false;
            _runStage2.Enabled = false;
            SetBusy(true);
            try
            {
                Log("开始阶段1，请不要关闭应用窗口。");
                await Task.Run(() =>
                {
                    var service = new Stage1Service(new ClosedXmlStage1ExcelGateway());
                    var report = service.Run(options, LogThreadSafe);
                    LogThreadSafe("阶段1完成。");
                    LogThreadSafe("输出台账：" + report.Output);
                    LogThreadSafe("报告：" + report.ReportPath);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
            }
            finally
            {
                SetBusy(false);
                _runStage1.Enabled = true;
                _cleanPower.Enabled = true;
                _runStage2.Enabled = true;
            }
        }

        private async Task RunCleanPowerAsync()
        {
            string rawDetailPath;
            string outputPath;
            try
            {
                rawDetailPath = _rawDetail.Text.Trim();
                outputPath = ResolvePowerOutputPath(rawDetailPath);
                _power.Text = outputPath;

                var message = "确认只清洗电量数据？" + Environment.NewLine
                    + "输出文件：" + outputPath;
                if (MessageBox.Show(this, message, "确认清洗电量", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
                return;
            }

            _runStage1.Enabled = false;
            _cleanPower.Enabled = false;
            _runStage2.Enabled = false;
            SetBusy(true);
            try
            {
                Log("开始清洗电量数据。");
                await Task.Run(() =>
                {
                    var service = new Stage1Service(new ClosedXmlStage1ExcelGateway());
                    var report = service.CleanPowerData(rawDetailPath, outputPath, LogThreadSafe);
                    LogThreadSafe("电量清洗完成。");
                    LogThreadSafe("电量处理表：" + report.OutputPath);
                    LogThreadSafe("客户数量：" + report.PowerRows + "，合计电量：" + report.MonthTotal.ToString("0.####"));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
            }
            finally
            {
                SetBusy(false);
                _runStage1.Enabled = true;
                _cleanPower.Enabled = true;
                _runStage2.Enabled = true;
            }
        }

        private async Task RunStage2Async()
        {
            Stage2Options options;
            try
            {
                options = CreateStage2Options();
                if (!ConfirmRun("阶段二", options.Month, options.OutputDirectory))
                {
                    return;
                }

                var service = new Stage2Service(new ClosedXmlStage1ExcelGateway());
                var preflight = service.Analyze(options);
                if (preflight.HasIssues && !ConfirmStage2Preflight(preflight))
                {
                    Log("已取消阶段2生成。");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
                return;
            }

            _runStage1.Enabled = false;
            _cleanPower.Enabled = false;
            _runStage2.Enabled = false;
            SetBusy(true);
            try
            {
                Log("开始阶段2，请确认台账和模板文件没有在 Excel 中打开。");
                await Task.Run(() =>
                {
                    var service = new Stage2Service(new ClosedXmlStage1ExcelGateway());
                    var report = service.Run(options, LogThreadSafe);
                    LogThreadSafe("阶段2完成。");
                    LogThreadSafe("汇总表：" + report.Summary);
                    LogThreadSafe("报告：" + report.ReportPath);
                    LogThreadSafe("代理费合计：" + report.ProxyTotal.ToString("0.####"));
                    LogThreadSafe("居间费合计：" + report.IntermediaryTotal.ToString("0.####"));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
            }
            finally
            {
                SetBusy(false);
                _runStage1.Enabled = true;
                _cleanPower.Enabled = true;
                _runStage2.Enabled = true;
            }
        }

        private Stage1Options CreateOptions()
        {
            var powerPath = _power.Text.Trim();
            var rawDetailPath = _rawDetail.Text.Trim();
            if (!string.IsNullOrWhiteSpace(rawDetailPath))
            {
                powerPath = ResolvePowerOutputPath(rawDetailPath);
                _power.Text = powerPath;
            }

            return new Stage1Options
            {
                Month = SelectedMonth(),
                BaseLedgerPath = _baseLedger.Text.Trim(),
                PowerPath = powerPath,
                RawDetailPath = rawDetailPath,
                ReferenceLedgerPath = _referenceLedger.Text.Trim(),
                OutputDirectory = _outputDir.Text.Trim(),
                CopyReferenceExisting = _copyReferenceExisting.Checked
            };
        }

        private string ResolvePowerOutputPath(string rawDetailPath)
        {
            if (string.IsNullOrWhiteSpace(rawDetailPath))
            {
                throw new InvalidOperationException("请选择原始零售侧明细。");
            }

            var outputDirectory = _outputDir.Text.Trim();
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                throw new InvalidOperationException("请选择结果输出文件夹。");
            }

            return Path.Combine(outputDirectory, "零售侧用户电量数据处理表.xlsx");
        }

        private Stage2Options CreateStage2Options()
        {
            return new Stage2Options
            {
                Month = SelectedMonth(),
                LedgerPath = _completedLedger.Text.Trim(),
                ProxyTemplateDirectory = _proxyTemplateDir.Text.Trim(),
                IntermediaryTemplateDirectory = _intermediaryTemplateDir.Text.Trim(),
                SummaryTemplatePath = _summaryTemplate.Text.Trim(),
                OutputDirectory = _outputDir.Text.Trim(),
                AllowMissingOwner = _allowMissingOwner.Checked
            };
        }

        private int SelectedMonth()
        {
            if (_month.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择结算月份。");
            }

            return _month.SelectedIndex + 2;
        }

        private bool ConfirmRun(string stageName, int month, string outputDirectory)
        {
            var message = "确认运行" + stageName + "？" + Environment.NewLine
                + "结算月份：2026年" + month + "月" + Environment.NewLine
                + "输出文件夹：" + outputDirectory;
            return MessageBox.Show(this, message, "确认运行", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK;
        }

        private bool ConfirmStage2Preflight(Stage2PreflightReport report)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "阶段二预检确认";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Size = new Size(760, 560);
                dialog.MinimumSize = new Size(680, 500);
                dialog.BackColor = WindowBackground;
                dialog.Font = Font;

                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    RowCount = 4,
                    ColumnCount = 1,
                    Padding = new Padding(18),
                    BackColor = WindowBackground
                };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                dialog.Controls.Add(layout);

                layout.Controls.Add(new Label
                {
                    Text = "发现需要确认的阶段二变化",
                    AutoSize = true,
                    ForeColor = MainText,
                    Font = new Font(Font.FontFamily, 14F, FontStyle.Bold),
                    Margin = new Padding(0, 0, 0, 8)
                }, 0, 0);

                layout.Controls.Add(new Label
                {
                    Text = "结算月份：2026年" + report.Month + "月；共发现 " + report.Issues.Count + " 条需要确认的变化。确认后仍会生成文件，并写入阶段二校验报告。",
                    AutoSize = true,
                    MaximumSize = new Size(700, 0),
                    ForeColor = MutedText,
                    Margin = new Padding(0, 0, 0, 12)
                }, 0, 1);

                var detail = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Text = BuildPreflightText(report),
                    BackColor = Color.White,
                    ForeColor = MainText,
                    BorderStyle = BorderStyle.FixedSingle,
                    Margin = new Padding(0, 0, 0, 14)
                };
                layout.Controls.Add(detail, 0, 2);

                var actions = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = true
                };
                var ok = new Button { Text = "继续生成并写报告", Width = 150, Height = 34, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Width = 90, Height = 34, DialogResult = DialogResult.Cancel, Margin = new Padding(8, 0, 0, 0) };
                StylePrimaryButton(ok);
                StyleSecondaryButton(cancel);
                actions.Controls.Add(ok);
                actions.Controls.Add(cancel);
                layout.Controls.Add(actions, 0, 3);

                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;
                return dialog.ShowDialog(this) == DialogResult.OK;
            }
        }

        private static string BuildPreflightText(Stage2PreflightReport report)
        {
            var builder = new StringBuilder();
            var grouped = report.Issues
                .GroupBy(issue => issue.Category)
                .OrderBy(group => group.Key);

            foreach (var group in grouped)
            {
                builder.AppendLine("【" + group.Key + "】");
                var index = 1;
                foreach (var issue in group)
                {
                    builder.AppendLine(index + ". [" + issue.Severity + "] " + issue.Message);
                    if (!string.IsNullOrWhiteSpace(issue.Kind))
                    {
                        builder.AppendLine("   类型：" + issue.Kind);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Owner) || !string.IsNullOrWhiteSpace(issue.Entity))
                    {
                        builder.AppendLine("   负责人/主体：" + issue.Owner + " / " + issue.Entity);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Customer))
                    {
                        builder.AppendLine("   客户：" + issue.Customer);
                    }
                    if (issue.LedgerRow > 0)
                    {
                        builder.AppendLine("   台账行：" + issue.LedgerRow);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.PreviousValue) || !string.IsNullOrWhiteSpace(issue.CurrentValue))
                    {
                        builder.AppendLine("   对比：" + issue.PreviousValue + "；" + issue.CurrentValue);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Suggestion))
                    {
                        builder.AppendLine("   建议：" + issue.Suggestion);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
                    {
                        builder.AppendLine("   上月模板：" + issue.TemplateFile);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.SheetName))
                    {
                        builder.AppendLine("   工作表：" + issue.SheetName);
                    }
                    builder.AppendLine();
                    index++;
                }
            }

            return builder.ToString();
        }

        private void LogThreadSafe(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(Log), message);
                return;
            }
            Log(message);
        }

        private void Log(string message)
        {
            _log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }

        private void SetBusy(bool busy)
        {
            _status.Text = busy ? "运行中" : "就绪";
            _status.ForeColor = busy ? Color.FromArgb(131, 82, 0) : Accent;
            _status.BackColor = busy ? Color.FromArgb(255, 243, 217) : Color.FromArgb(230, 246, 247);
        }
    }
}
