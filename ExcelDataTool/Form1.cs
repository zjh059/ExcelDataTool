using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

//11：03
namespace ExcelDataTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // ★ 新版 EPPlus 的免费许可声明方式
            ExcelPackage.License.SetNonCommercialPersonal("团队协作使用");

            // ★ 初始化下拉菜单的选项
            cmbTaskSelect.Items.Add("大孟: 2025年2月 - 2025年5月");
            cmbTaskSelect.Items.Add("zjh: 2025年6月 - 2025年12月");
            cmbTaskSelect.Items.Add("Lucky: 2026年1月 - 2026年3月");

            // 默认选中第2项（索引为1，也就是你的任务）
            cmbTaskSelect.SelectedIndex = 1;
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel 文件|*.xlsx;*.xls";
                ofd.Title = "请选择你的数据总表";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                    AppendLog($"[准备就绪] 已选中文件：{ofd.FileName}");
                }
            }
        }

        private async void btnProcess_Click(object sender, EventArgs e)
        {
            string inputFilePath = txtFilePath.Text;
            if (string.IsNullOrWhiteSpace(inputFilePath) || !File.Exists(inputFilePath))
            {
                MessageBox.Show("请先选择正确的文件路径！");
                return;
            }

            // ★ 核心逻辑：根据下拉菜单的选择，决定起止时间和输出文件名
            DateTime startDate = DateTime.MinValue;
            DateTime endDate = DateTime.MinValue;
            string fileNameSuffix = "";

            if (cmbTaskSelect.SelectedIndex == 0) // 同事A
            {
                startDate = new DateTime(2025, 2, 1);
                endDate = new DateTime(2025, 5, 31);
                fileNameSuffix = "2025年2月至5月";
            }
            else if (cmbTaskSelect.SelectedIndex == 1) // 你的任务
            {
                startDate = new DateTime(2025, 6, 1);
                endDate = new DateTime(2025, 12, 31);
                fileNameSuffix = "2025年6月至12月";
            }
            else if (cmbTaskSelect.SelectedIndex == 2) // 同事B (Lucky)
            {
                startDate = new DateTime(2026, 1, 1);
                endDate = new DateTime(2026, 3, 31);
                fileNameSuffix = "2026年1月至3月";
            }

            // 动态生成防混淆的文件名
            string outputFilePath = Path.Combine(Path.GetDirectoryName(inputFilePath), $"分类处理完成_{fileNameSuffix}.xlsx");

            btnSelectFile.Enabled = false;
            btnProcess.Enabled = false;
            cmbTaskSelect.Enabled = false; // 处理期间不准修改任务

            AppendLog("=========================================");
            AppendLog($"🚀 启动任务！当前目标：提取 {startDate:yyyy年MM月} 到 {endDate:yyyy年MM月} 的数据...");

            // 把动态的起止时间传给处理函数
            await Task.Run(() => ProcessExcelData(inputFilePath, outputFilePath, startDate, endDate));

            btnSelectFile.Enabled = true;
            btnProcess.Enabled = true;
            cmbTaskSelect.Enabled = true;
        }

        // 核心数据处理函数
        private void ProcessExcelData(string inputFilePath, string outputFilePath, DateTime startDate, DateTime endDate)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
                {
                    AppendLog($"✅ 成功打开表格，准备扫描所有 Sheet...");

                    var headers = new List<string>();
                    var dailyData = new Dictionary<DateTime, List<List<object>>>();
                    int validCount = 0;

                    // ★ 核心改动：遍历 Excel 中的每一个 Sheet
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        // 如果遇到完全空白的 Sheet，直接跳过，防止报错
                        if (worksheet.Dimension == null) continue;

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        AppendLog($"🔍 正在扫描表单：[{worksheet.Name}]，共 {rowCount} 行...");

                        // 只在第一次循环时（headers为空）提取表头，假设所有 Sheet 结构一样
                        if (headers.Count == 0)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                headers.Add(worksheet.Cells[1, col].Text);
                            }
                        }

                        // 遍历当前 Sheet 的数据（从第2行开始，跳过表头）
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var dateValue = worksheet.Cells[row, 7].Value;
                            DateTime rowDate = DateTime.MinValue;
                            bool isValidDate = false;

                            if (dateValue is DateTime dt) { rowDate = dt.Date; isValidDate = true; }
                            else if (dateValue is double d) { rowDate = DateTime.FromOADate(d).Date; isValidDate = true; }
                            else if (DateTime.TryParse(dateValue?.ToString(), out DateTime parsedDt)) { rowDate = parsedDt.Date; isValidDate = true; }

                            if (isValidDate && rowDate >= startDate && rowDate <= endDate)
                            {
                                var rowData = new List<object>();
                                for (int col = 1; col <= colCount; col++)
                                {
                                    rowData.Add(worksheet.Cells[row, col].Value);
                                }

                                if (!dailyData.ContainsKey(rowDate))
                                {
                                    dailyData[rowDate] = new List<List<object>>();
                                }
                                dailyData[rowDate].Add(rowData);
                                validCount++;
                            }

                            // 进度提示：每 50000 行打印一次
                            if (row % 50000 == 0)
                            {
                                AppendLog($"...[{worksheet.Name}] 正在扫描第 {row} 行...");
                            }
                        }
                    }

                    AppendLog($"🎯 筛选完毕！所有表单中符合条件的数据总计：{validCount} 条。开始按规则写入新表...");

                    // 开始写入新文件
                    using (var outPackage = new ExcelPackage(new FileInfo(outputFilePath)))
                    {
                        var smallDataList = new List<List<object>>();

                        foreach (var kvp in dailyData.OrderBy(x => x.Key))
                        {
                            DateTime date = kvp.Key;
                            var rows = kvp.Value;
                            string dateStr = date.ToString("yyyy-MM-dd");

                            if (rows.Count > 1000)
                            {
                                AppendLog($"✂️ [{dateStr}] 当日数据 {rows.Count} 条(>1000)，已建立专属 Sheet。");
                                var newSheet = outPackage.Workbook.Worksheets.Add(dateStr);
                                WriteDataToSheet(newSheet, headers, rows);
                            }
                            else
                            {
                                AppendLog($"📦 [{dateStr}] 当日数据 {rows.Count} 条(<=1000)，放入合并等待区。");
                                smallDataList.AddRange(rows);
                            }
                        }

                        if (smallDataList.Count > 0)
                        {
                            AppendLog($"🧹 将合并区内共 {smallDataList.Count} 条数据生成【小于1000条汇总】Sheet...");
                            var summarySheet = outPackage.Workbook.Worksheets.Add("小于1000条汇总");
                            WriteDataToSheet(summarySheet, headers, smallDataList);
                        }

                        AppendLog("💾 正在保存新文件至硬盘...");
                        outPackage.Save();
                    }

                    AppendLog("=========================================");
                    AppendLog($"🎉 大功告成！文件绝对不会混淆，新文件位置：");
                    AppendLog(outputFilePath);
                }
            }
            catch (Exception ex)
            {
                AppendLog($"❌ 遇到错误了：{ex.Message}");
            }
        }

        private void WriteDataToSheet(ExcelWorksheet sheet, List<string> headers, List<List<object>> data)
        {
            for (int i = 0; i < headers.Count; i++) { sheet.Cells[1, i + 1].Value = headers[i]; }
            for (int row = 0; row < data.Count; row++)
            {
                for (int col = 0; col < headers.Count; col++) { sheet.Cells[row + 2, col + 1].Value = data[row][col]; }
            }
            sheet.Column(7).Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
        }

        private void AppendLog(string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendLog), msg);
            }
            else
            {
                rtbLogs.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\r\n");
                rtbLogs.ScrollToCaret();
            }
        }
    }
}