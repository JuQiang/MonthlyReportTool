using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.TeamProject;
using System.Diagnostics;

namespace MonthlyReportTool.API.Office.Excel
{
    public class BugSheet : ExcelSheetBase, IExcelSheet
    {
        private List<List<BugEntity>> bugList;
        private ExcelInterop.Worksheet sheet;
        private ProjectEntity project;
        public BugSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            this.project = project;
            this.bugList = Bug.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));
            BuildTitle();
            BuildSubTitle();

            BuildDescription();

            BuildSummaryTable();

            int startRow = BuildFixedRateTable(14, new List<List<BugEntity>>() { this.bugList[0], this.bugList[2], this.bugList[1]});
            startRow = BuildReasonTable(startRow, this.bugList[3]);
            startRow = BuildNoneTable(startRow, this.bugList[4]);
            startRow = BuildCodeReviewTable(startRow, this.bugList[5]);
            startRow = BuildAddedTable(startRow, this.bugList[0]);
            startRow = BuildNotResolvedTable(startRow, this.bugList[2]);

            sheet.Cells[1, "A"] = "";
        }

        private int BuildCodeReviewTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代代码评审发现问题统计", "说明：", "B", "L",
                new List<string>() { "BugID", "关键应用", "模块","缺陷发现方式", "问题类别", "严重级别", "Bug标题", "指派给", "发现人" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F", "G,G", "H,J","K,K", "L,L" },
                list.Count);

            startRow += 3;
            object[,] arr = new object[list.Count, 12];
            for (int i = 0; i < list.Count; i++)
            {
                arr[i, 0] = list[i].Id.ToString();
                arr[i, 1] = list[i].KeyApplication;
                arr[i, 2] = list[i].ModulesName;
                arr[i, 3] = list[i].DetectionMode;
                arr[i, 4] = list[i].Type;
                arr[i, 5] = list[i].Severity;
                arr[i, 6] = list[i].Title;
                arr[i, 9] = Utility.GetPersonName(list[i].AssignedTo);
                arr[i, 10] = Utility.GetPersonName(list[i].DiscoveryUser);
            }

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "L"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "B"]]);

            return nextRow;
        }

        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "I", "Bug统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "I"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "Bug新增及处理情况统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 12;

        }

        private void BuildDescription()
        {
            int row = 5;
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "E"]];
            Utility.AddNativieResource(titleRange);
            titleRange.RowHeight = 80;
            titleRange.Merge();
            sheet.Cells[row, "B"] = "说明：\r\n" +
                                    "      Bug修复率 = 本迭代修复数 /（本迭代遗留数 + 本迭代修复数）\r\n" +
                                    "      Bug遗留率 = 本迭代遗留数）/（本迭代遗留数 + 本迭代修复数）\r\n" +
                                    "      1、2级占比 = (1、2级问题数)/ 合计里面的数\r\n" +
                                    "      平均Bug修复耗时：（关闭日期 - 创建日期）/ 总修复数";

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;

            ExcelInterop.Range titleRange2 = sheet.Range[sheet.Cells[row, "F"], sheet.Cells[row, "I"]];
            titleRange2.Merge();
            Utility.AddNativieResource(titleRange2);
            titleRange.Merge();
            sheet.Cells[row, "F"] = "\r\n" +
                                    "本迭代新增数：本迭代新登记的Bug数\r\n" +
                                    "本迭代修复数：本迭代修复的所有的Bug数（包括非本迭代登记的Bug）\r\n" +
                                    "本迭代遗留数：本迭代结束后，遗留未修复的所有Bug数（包括非本迭代登记的Bug）";
        }

        private void BuildSummaryTable()
        {
            int start = 6;
            string[,] cols = new string[,]
                        {
                { "", "1 - 严重", "2 - 高", "3 - 中","4 - 低","5 - 无（建议）","合计","1、2级占比"},
                { "本迭代新增数", "", "", "","", "", "",""},
                { "本迭代遗留数", "", "", "","", "", "",""},
                { "本迭代修复数", "", "", "","", "", "",""},
                { "不予处理/不是错误数", "", "", "","", "", "",""},
                { "代码评审问题数", "", "", "","", "", "",""},
                { "", "", "", "","", "", "",""},
                        };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","E"),
                Tuple.Create<string,string>("F","F"),
                Tuple.Create<string,string>("G","G"),
                Tuple.Create<string,string>("H","H"),
                Tuple.Create<string,string>("I","I"),

            };
            for (int row = 0; row < cols.GetLength(0); row++)
            {
                for (int col = 0; col < colsname.Count; col++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[start + row, colsname[col].Item1], sheet.Cells[start + row, colsname[col].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[start + row, colsname[col].Item1] = cols[row, col];

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[start, "B"], sheet.Cells[start, "I"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;


            sheet.Cells[7, "H"] = "=SUM(C7:G7)"; sheet.Cells[8, "H"] = "=SUM(C8:G8)"; sheet.Cells[9, "H"] = "=SUM(C9:G9)"; sheet.Cells[10, "H"] = "=SUM(C10:G10)"; sheet.Cells[11, "H"] = "=SUM(C11:G11)";
            sheet.Cells[7, "I"] = "=(C7+D7)/H7"; sheet.Cells[8, "I"] = "=(C8+D8)/H8"; sheet.Cells[9, "I"] = "=(C9+D9)/H9"; 

            sheet.Cells[12, "B"] = "Bug修复率";
            sheet.Cells[12, "C"] = "=(H9)/(H8+H9)";
            sheet.Cells[12, "E"] = "Bug遗留率";
            sheet.Cells[12, "F"] = "=1-C12";

            Utility.SetupSheetPercentFormat(sheet, sheet.get_Range("C12:C12,F12:F12,I7:I9"));

            List<List<BugEntity>> list = new List<List<BugEntity>>();
            list.Add(this.bugList[0]);
            list.Add(this.bugList[2]);
            list.Add(this.bugList[1]);
            list.Add(this.bugList[4]);
            list.Add(this.bugList[5]);

            for (int i = 0; i < list.Count; i++)
            {
                sheet.Cells[7+i,"C"] = list[i].Where(bug => bug.Severity == "1 - 严重").Count();
                sheet.Cells[7 + i, "D"] = list[i].Where(bug => bug.Severity == "2 - 高").Count();
                sheet.Cells[7 + i, "E"] = list[i].Where(bug => bug.Severity == "3 - 中").Count();
                sheet.Cells[7 + i, "F"] = list[i].Where(bug => bug.Severity == "4 - 低").Count();
                sheet.Cells[7 + i, "G"] = list[i].Where(bug => bug.Severity == "5 - 无").Count();
            }
        }

        private int BuildFixedRateTable(int startRow, List<List<BugEntity>> list)
        {
            List<string> members = new List<string>();

            for (int i = 0; i < list.Count; i++)
            {
                foreach (var bug in list[i].GroupBy(bug => bug.AssignedTo))
                {
                    if (members.Contains(bug.Key)) continue;
                    //if (bug.Key.Trim().Length < 1) members.Add("（未指定）");
                    members.Add(bug.Key);
                }
            }
            
            int nextRow = Utility.BuildFormalTable(sheet, startRow, "按人员统计的修复率", "说明：", "B", "F",
                new List<string>() { "姓名", "本迭代新增数", "本迭代遗留数", "本迭代修复数", "Bug修复率" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F" },
                members.Count
                );

            startRow += 3;
            for (int i = 0; i < members.Count; i++)
            {
                sheet.Cells[startRow + i, "F"] = String.Format("=E{0}/(D{0}+E{0})", startRow + i);

                if (members[i].Trim().Length < 1)
                {
                    sheet.Cells[startRow + i, "B"] = "（未指定）";
                }
                else
                {
                    sheet.Cells[startRow + i, "B"] = members[i].Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                }
                sheet.Cells[startRow + i, "C"] = list[0].Where(bug => bug.AssignedTo == members[i]).Count();
                sheet.Cells[startRow + i, "D"] = list[1].Where(bug => bug.AssignedTo == members[i]).Count();
                sheet.Cells[startRow + i, "E"] = list[2].Where(bug => bug.AssignedTo == members[i]).Count();
            }

            sheet.Cells[startRow + members.Count, "B"] = "合计";
            sheet.Cells[startRow + members.Count, "C"] = String.Format("=SUM(C{0}:C{1})", startRow, startRow + members.Count - 1);
            sheet.Cells[startRow + members.Count, "D"] = String.Format("=SUM(D{0}:D{1})", startRow, startRow + members.Count - 1);
            sheet.Cells[startRow + members.Count, "E"] = String.Format("=SUM(E{0}:E{1})", startRow, startRow + members.Count - 1);
            sheet.Cells[startRow + members.Count, "F"] = String.Format("=E{0}/(D{0}+E{0})", startRow + members.Count);

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + members.Count, "B"]]);
            Utility.SetCellBorder(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + members.Count, "B"]]);
            Utility.SetupSheetPercentFormat(sheet,sheet.Range[sheet.Cells[startRow, "F"],sheet.Cells[startRow + members.Count, "F"]]);

            AddBugChart(sheet,
                startRow-2, "G", startRow + members.Count-1, "L",
                String.Format("B{0}:B{1},C{0}:C{1}", startRow-1/*包含标题列和标题行*/, startRow + members.Count - 1),
                "开发人员新增bug数");

            AddBugChart(sheet,
                startRow-2, "O", startRow + members.Count - 1, "T",
                String.Format("B{0}:B{1},F{0}:F{1}", startRow-1, startRow + members.Count - 1),
                "开发人员修复率");
           
            return nextRow + 1;
        }

        private int BuildReasonTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "Bug产生原因分析", "说明：主要针对严重级别为1、2级的Bug进行原因分析", "B", "N",
                new List<string>() { "BugID","关键应用","模块", "问题类别", "严重级别", "Bug标题", "指派给", "发现人", "原因分析"},
                new List<string>() { "B,B", "C,C", "D,D","E,E","F,F", "G,I", "J,J", "K,K", "L,N" },
                list.Count);

            startRow += 3;

            object[,] arr = new object[list.Count, 14];
            for (int i = 0; i < list.Count; i++)
            {
                arr[i, 0] = list[i].Id;
                arr[i, 1] = list[i].KeyApplication;
                arr[i, 2] = list[i].ModulesName;
                arr[i, 3] = list[i].Type;
                arr[i, 4] = list[i].Severity;
                arr[i, 5] = list[i].Title;
                arr[i, 8] = Utility.GetPersonName(list[i].AssignedTo);
                arr[i, 9] = Utility.GetPersonName(list[i].DiscoveryUser);
            }

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "L"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "B"]]);

            Utility.SetCellRedColor(sheet.Cells[startRow - 1, "L"]);

            return nextRow;
        }
        private int BuildAddedTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代新增Bug数", "说明：", "B", "L",
                new List<string>() { "BugID", "关键应用", "模块", "问题类别", "严重级别", "Bug标题", "指派给", "发现人" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E","F,F","G,J", "K,K", "L,L" },
                list.Count);

            startRow += 3;
            
            object[,] arr = new object[list.Count, 12+1];
            for (int i = 0; i < list.Count; i++)
            {
                arr[i, 0] = list[i].Id;
                arr[i, 1] = list[i].KeyApplication;
                arr[i, 2] = list[i].ModulesName;
                arr[i, 3] = list[i].Type;
                arr[i, 4] = list[i].Severity;
                arr[i, 5] = list[i].Title;
                arr[i, 9] = Utility.GetPersonName(list[i].AssignedTo);
                arr[i, 10] = Utility.GetPersonName(list[i].DiscoveryUser);
            }
            

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "L"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "B"]]);

            return nextRow;
        }
        private int BuildNoneTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代处理的不是错误/不予处理Bug分析", "说明：", "B", "N",
                new List<string>() { "BugID", "关键应用", "模块", "关闭原因", "问题类别", "严重级别", "Bug标题","指派给", "不是错误/不予处理分析" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E","F,F", "G,G", "H,J","K,K","L,N" },
                list.Count);

            startRow += 3;
            object[,] arr = new object[list.Count, 14];
            for (int i = 0; i < list.Count; i++)
            {
                arr[i, 0] = list[i].Id;
                arr[i, 1] = list[i].KeyApplication;
                arr[i, 2] = list[i].ModulesName;
                arr[i, 3] = list[i].ResolvedReason;
                arr[i, 4] = list[i].Type;
                arr[i, 5] = list[i].Severity;
                arr[i, 6] = list[i].Title;
                arr[i, 9] = Utility.GetPersonName(list[i].AssignedTo);
                arr[i, 10] = "";
            }


            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "N"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "B"]]);
            Utility.SetCellRedColor(sheet.Cells[startRow - 1, "L"]);

            return nextRow;
        }

        private int BuildNotResolvedTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代遗留Bug数", "说明：", "B", "L",
                new List<string>() { "BugID", "关键应用", "模块", "问题类别", "严重级别", "Bug标题", "指派给", "发现人" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,F", "G,J", "K,K", "L,L" },
                list.Count);

            startRow += 3;
            object[,] arr = new object[list.Count, 12 + 1];
            for (int i = 0; i < list.Count; i++)
            {
                arr[i, 0] = list[i].Id;//这里是col对应的数字，不能按照0开始算
                arr[i, 1] = list[i].KeyApplication;
                arr[i, 2] = list[i].ModulesName;
                arr[i, 3] = list[i].Type;
                arr[i, 4] = list[i].Severity;
                arr[i, 5] = list[i].Title;
                arr[i, 9] = Utility.GetPersonName(list[i].AssignedTo);
                arr[i, 10] = Utility.GetPersonName(list[i].DiscoveryUser);
            }


            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "L"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + list.Count - 1, "B"]]);

            return nextRow;
        }
        private void AddBugChart(
            ExcelInterop.Worksheet sheet, 
            int chartStartRow, string chartStartCol, int chartEndRow, string chartEndCol,
            string bugDataSource,string chartTitle) {

            ExcelInterop.Range bugChartRange = sheet.Range[sheet.Cells[chartStartRow, chartStartCol], sheet.Cells[chartEndRow, chartEndCol]];

            ExcelInterop.ChartObjects charts = sheet.ChartObjects(Type.Missing) as ExcelInterop.ChartObjects;
            Utility.AddNativieResource(charts);

            ExcelInterop.ChartObject bugChartObject = charts.Add(0, 0, bugChartRange.Width, bugChartRange.Height);
            Utility.AddNativieResource(bugChartObject);
            ExcelInterop.Chart bugChart = bugChartObject.Chart;//设置图表数据区域。
            Utility.AddNativieResource(bugChart);

            ExcelInterop.Range datasource = sheet.get_Range(bugDataSource);//不是："B14:B25","F14:F25"
            Utility.AddNativieResource(datasource);
            bugChart.SetSourceData(datasource);
            bugChart.ChartType = ExcelInterop.XlChartType.xlColumnClustered;
            //bugChart.ChartWizard(datasource, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, chartTitle, "", "", Type.Missing);
            bugChart.ApplyDataLabels();//图形上面显示具体的值
            
            //将图表移到数据区域之下。
            bugChartObject.Left = Convert.ToDouble(bugChartRange.Left)+20;
            bugChartObject.Top = Convert.ToDouble(bugChartRange.Top) + 20;

            bugChartObject.Locked = false;
            bugChartObject.Select();
            bugChartObject.Activate();
        }
    }
}
