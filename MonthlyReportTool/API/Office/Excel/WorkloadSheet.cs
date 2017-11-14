using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class WorkloadSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public WorkloadSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(string project)
        {
            BuildTitle();
            BuildSubTitle();
            BuildDescription();
            BuildSummaryTable();

            BuildDevelopmentTitle();
            BuildDevelopmentDescription();

            int startRow = BuildDevelopmentTableTitle();

        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "AG", "工作量统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "团队整体工作量分析";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 12;

        }
        private void BuildDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "AG"]];
            Utility.AddNativieResource(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明：容量投入总工时：迭代成员迭代规划容量×迭代天数之和计算";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;
        }
        private int BuildSummaryTable()
        {
            int featuresCount = 1;
            string[] cols = new string[] { "团队成员标准工时\r\n（迭代天数*人数×8）", "实际投入总工时", "工作\r\n饱和度", "研发投入总工时", "研发投入\r\n占比", "容量投入总工时\r\n（迭代天数×人数×容量）", "评估总工时\r\n（迭代任务的评估工时）", "剩余工时" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","E"),
                Tuple.Create<string,string>("F","H"),
                Tuple.Create<string,string>("I","J"),
                Tuple.Create<string,string>("K","L"),
                Tuple.Create<string,string>("M","N"),
                Tuple.Create<string,string>("O","Q"),
                Tuple.Create<string,string>("R","T"),
                Tuple.Create<string,string>("U","W"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[6, colsname[i].Item1], sheet.Cells[6, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.Merge();
                sheet.Cells[6, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "W"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 50;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[7 + i, colsname[j].Item1], sheet.Cells[7 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[7 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 7 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return 7 + featuresCount + 2;
        }

        private void BuildDevelopmentTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[9, "B"], sheet.Cells[9, "AG"]];
            Utility.AddNativieResource(range);
            range.RowHeight = 20;
            range.Merge();
            sheet.Cells[9, "B"] = "开发工作量统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
        }

        private void BuildDevelopmentDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[10, "B"], sheet.Cells[10, "K"]];
            Utility.AddNativieResource(titleRange);
            titleRange.RowHeight = 120;
            titleRange.Merge();
            sheet.Cells[10, "B"] = "说明：包括测试人员的工作量统计\r\n" +
           "        开发人员按照团队成员的研发工作量占比排序\r\n" +
           "        测试人员按照团队成员的Bug产出率排序\r\n" +
           "        各明细工作量是实际填写的工作日志某类的工作量，和工作日志中分类一致\r\n" +
           "        Bug数：对于开发人员是本迭代被测试出的Bug总数\r\n" +
           "                 对于测试人员是本迭代测试出的Bug总数";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;

            var tmpchar2 = titleRange.Characters[30, 45];

            var tmpfont2 = tmpchar2.Font;
            Utility.AddNativieResource(tmpchar2);
            Utility.AddNativieResource(tmpfont2);
            tmpfont2.Color = System.Drawing.Color.Red.ToArgb();

            ExcelInterop.Range titleRange2 = sheet.Range[sheet.Cells[10, "L"], sheet.Cells[10, "W"]];
            titleRange2.Merge();
            Utility.AddNativieResource(titleRange2);
            titleRange.Merge();
            sheet.Cells[10, "L"] = "      标准工作量：迭代天数*8\r\n" +
                                   "      实际投入工作量：不包含请假实际的工作量\r\n" +
                                   "      实际饱和度：实际投入工作量 / 标准工作量\r\n" +
                                   "      Bug产出率：bug数 / 实际投入工作量";
        }

        private int BuildDevelopmentTableTitle()
        {
            string[,] cols = new string[,]
                        {
                { "团队成员", "标准\r\n工作量", "实际投入\r\n工作量", "请假","实际\r\n饱和度","bug数","bug产出率（个/天）","研发工作量占比","研发","","","","","","","","","","","","管理","","运维","","文档","","学习交流","","售前/推广","","其他\r\n排除请假",""},
                { "", "", "", "", "", "", "", "", "开发","","需求","","设计","","测试设计","","测试执行","","其他","","","","","","","","","","","","",""},
                { "", "", "", "", "", "", "", "", "工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比"},
                        };

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < cols.GetLength(1); col++)
                {
                    sheet.Cells[row + 11, col + 2] = cols[row, col];

                }
            }

            List<Tuple<int, string, int, string>> allMergedCells = new List<Tuple<int, string, int, string>>()
            {
                Tuple.Create<int, string, int, string>(11,"B",13,"B"),Tuple.Create<int, string, int, string>(11,"C",13,"C"),
                Tuple.Create<int, string, int, string>(11,"D",13,"D"),Tuple.Create<int, string, int, string>(11,"E",13,"E"),
                Tuple.Create<int, string, int, string>(11,"F",13,"F"),Tuple.Create<int, string, int, string>(11,"G",13,"G"),
                Tuple.Create<int, string, int, string>(11,"H",13,"H"),Tuple.Create<int, string, int, string>(11,"I",13,"I"),

                Tuple.Create<int, string, int, string>(11,"J",11,"U"),

                Tuple.Create<int, string, int, string>(11,"V",12,"W"),Tuple.Create<int, string, int, string>(11,"X",12,"Y"),
                Tuple.Create<int, string, int, string>(11,"Z",12,"AA"),Tuple.Create<int, string, int, string>(11,"AB",12,"AC"),
                Tuple.Create<int, string, int, string>(11,"AD",12,"AE"),Tuple.Create<int, string, int, string>(11,"AF",12,"AG"),

                Tuple.Create<int, string, int, string>(12,"J",12,"K"),Tuple.Create<int, string, int, string>(12,"L",12,"M"),
                Tuple.Create<int, string, int, string>(12,"N",12,"O"),Tuple.Create<int, string, int, string>(12,"P",12,"Q"),
                Tuple.Create<int, string, int, string>(12,"R",12,"S"),Tuple.Create<int, string, int, string>(12,"T",12,"U")
            };

            foreach (var tuple in allMergedCells)
            {
                ExcelInterop.Range range = sheet.Range[sheet.Cells[tuple.Item1, tuple.Item2], sheet.Cells[tuple.Item3, tuple.Item4]];
                Utility.AddNativieResource(range);
                range.Merge();
            }

            Utility.BuildFormalTableHeader(sheet, 11, "B", 13, "AG");

            int devWorkloadCount = 7;
            for (int i = 0; i < devWorkloadCount; i++)
            {
                for (int j = 2; j < 34; j++)
                {
                    sheet.Cells[14 + i, j] = "开发";
                }
            }

            for (int i = 0; i < devWorkloadCount; i++)
            {
                sheet.Cells[14 + i, "B"] = String.Format("开发人{0:d3}", i + 1);
            }

            Random r = new Random();
            for (int i = 0; i < devWorkloadCount; i++)
            {
                sheet.Cells[14 + i, "F"] = String.Format("{0}%", r.Next(150) + 1);
            }

            ExcelInterop.Range devRange = sheet.Range[sheet.Cells[14, "B"], sheet.Cells[14 + devWorkloadCount - 1, "AG"]];
            devRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignRight;
            devRange.WrapText = true;

            var borderDevRange = devRange.Borders;
            Utility.AddNativieResource(borderDevRange);
            borderDevRange.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            ExcelInterop.Range devRange2 = sheet.Range[sheet.Cells[14, "B"], sheet.Cells[14 + devWorkloadCount - 1, "B"]];
            Utility.AddNativieResource(devRange2);
            devRange2.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            int testWorkloadCount = 3;
            for (int i = 0; i < testWorkloadCount; i++)
            {
                for (int j = 2; j < 34; j++)
                {
                    sheet.Cells[14 + devWorkloadCount + 2 + i, j] = "测试";
                }
            }

            for (int i = 0; i < testWorkloadCount; i++)
            {
                sheet.Cells[14 + devWorkloadCount + 2 + i, "B"] = String.Format("测试人{0:d3}", i + 1);
            }

            r = new Random();
            for (int i = 0; i < testWorkloadCount; i++)
            {
                sheet.Cells[14 + devWorkloadCount + 2 + i, "F"] = String.Format("{0}%", r.Next(150) + 1);
                sheet.Cells[14 + devWorkloadCount + 2 + i, "G"] = r.Next(50) + 1;
            }

            ExcelInterop.Range testRange = sheet.Range[sheet.Cells[14 + devWorkloadCount + 2, "B"], sheet.Cells[14 + devWorkloadCount + 2 + testWorkloadCount, "AG"]];
            testRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignRight;
            testRange.WrapText = true;

            var borderTestRange = testRange.Borders;
            Utility.AddNativieResource(borderTestRange);
            borderTestRange.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            sheet.Cells[14 + devWorkloadCount + 2 + testWorkloadCount, "B"] = "合计";

            int chartstart = 14 + devWorkloadCount + 2 + testWorkloadCount + 2;

            ExcelInterop.Range workloadChartRange = sheet.Range[sheet.Cells[chartstart, "B"], sheet.Cells[chartstart + 10, "AG"]];

            ExcelInterop.ChartObjects charts = sheet.ChartObjects(Type.Missing) as ExcelInterop.ChartObjects;
            Utility.AddNativieResource(charts);

            ExcelInterop.ChartObject workloadChartObject = charts.Add(0, 0, workloadChartRange.Width, workloadChartRange.Height);
            Utility.AddNativieResource(workloadChartObject);
            ExcelInterop.Chart workloadChart = workloadChartObject.Chart;//设置图表数据区域。
            Utility.AddNativieResource(workloadChart);

            //=工作量统计!$B$14:$B$33,工作量统计!$F$14:$F$33
            ExcelInterop.Range datasource = sheet.get_Range(String.Format("B14:B{0},F14:F{0}", 14 + devWorkloadCount + 2 + testWorkloadCount - 1));//不是："B14:B25","F14:F25"
            //ExcelInterop.Range datasource = sheet.get_Range("B14:B25,F14:F25");//不是："B14:B25","F14:F25"
            Utility.AddNativieResource(datasource);
            workloadChart.ChartWizard(datasource, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, "工作量饱和度", "人员", "工作量", Type.Missing);
            workloadChart.ApplyDataLabels();//图形上面显示具体的值
            //将图表移到数据区域之下。
            workloadChartObject.Left = Convert.ToDouble(workloadChartRange.Left);
            workloadChartObject.Top = Convert.ToDouble(workloadChartRange.Top) + 20;

            chartstart += 12;
            ExcelInterop.Range bugChartRange = sheet.Range[sheet.Cells[chartstart, "B"], sheet.Cells[chartstart + 10, "K"]];
            ExcelInterop.ChartObject bugChartObject = charts.Add(0, 0, bugChartRange.Width, bugChartRange.Height);
            Utility.AddNativieResource(bugChartObject);
            ExcelInterop.Chart bugChart = bugChartObject.Chart;//设置图表数据区域。
            Utility.AddNativieResource(bugChart);

            //=工作量统计!$B$14:$B$33,工作量统计!$F$14:$F$33
            ExcelInterop.Range datasource2 = sheet.get_Range(String.Format("B{0}:B{1},G{0}:G{1}", 14 + devWorkloadCount + 2 - 1, 14 + devWorkloadCount + 2 + testWorkloadCount - 1));//不是："B14:B25","F14:F25"
            //ExcelInterop.Range datasource = sheet.get_Range("B14:B25,F14:F25");//不是："B14:B25","F14:F25"
            Utility.AddNativieResource(datasource2);
            bugChart.ChartWizard(datasource2, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, "测试人员BUG数", "人员", "BUG数", Type.Missing);
            bugChart.ApplyDataLabels();//图形上面显示具体的值
            //将图表移到数据区域之下。
            bugChartObject.Left = Convert.ToDouble(bugChartRange.Left);
            bugChartObject.Top = Convert.ToDouble(bugChartRange.Top) + 20;

            return 14 + devWorkloadCount + testWorkloadCount + 2 + 2 + 10 + 10;
        }
    }
}
