using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.TeamProject;

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
            startRow = BuildReasonTable(startRow, this.bugList[0]);
            startRow = BuildNoneTable(startRow, this.bugList[4]);
            startRow = BuildCodeReviewTable(startRow, this.bugList[5]);
            startRow = BuildAddedTable(startRow, this.bugList[0]);
        }

        private int BuildCodeReviewTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代代码评审发现问题统计", "说明：", "B", "H",
                new List<string>() { "BugID", "缺陷发现方式", "问题类别", "严重级别", "Bug标题"},
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,H"},
                list.Count);

            startRow += 3;
            for (int i = 0; i < list.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = list[i].Id;
                sheet.Cells[startRow + i, "C"] = list[i].DetectionMode;
                sheet.Cells[startRow + i, "D"] = list[i].Type;
                sheet.Cells[startRow + i, "E"] = list[i].Severity;
                sheet.Cells[startRow + i, "F"] = list[i].Title;
            }

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

            Utility.SetupSheetPercentFormat(sheet, 12, "C", 12, "C");
            Utility.SetupSheetPercentFormat(sheet, 12, "F", 12, "F");
            Utility.SetupSheetPercentFormat(sheet, 7, "I", 9, "I");

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
                sheet.Cells[7 + i, "G"] = list[i].Where(bug => bug.Severity == "5 - 无（建议）").Count();
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

            Utility.SetupSheetPercentFormat(sheet, startRow, "F", startRow + members.Count, "F");


            return nextRow + 1;
        }

        private int BuildReasonTable(int startRow, List<BugEntity> list)
        {
            var bugs = list.Where(bug => bug.Severity.StartsWith("1") || bug.Severity.StartsWith("2")).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "Bug产生原因分析", "说明：主要针对严重级别为1、2级的Bug进行原因分析", "B", "L",
                new List<string>() { "BugID", "问题类别", "严重级别", "Bug标题", "原因分析", "指派给", "测试人员" },
                new List<string>() { "B,B", "C,C", "D,D", "E,H", "I,J", "K,K", "L,L" },
                bugs.Count);

            startRow += 3;
            for (int i = 0; i < bugs.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = bugs[i].Id;
                sheet.Cells[startRow + i, "C"] = bugs[i].Type;
                sheet.Cells[startRow + i, "D"] = bugs[i].Severity;
                sheet.Cells[startRow + i, "E"] = bugs[i].Title;
                sheet.Cells[startRow + i, "I"] = "";
                sheet.Cells[startRow + i, "K"] = Utility.GetPersonName(bugs[i].AssignedTo);
                sheet.Cells[startRow + i, "L"] = Utility.GetPersonName(bugs[i].TestResponsibleMan);
            }

            return nextRow;
        }
        private int BuildAddedTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代新增Bug数", "说明：", "B", "J",
                new List<string>() { "BugID", "问题类别", "严重级别", "Bug标题", "指派给", "发现人" },
                new List<string>() { "B,B", "C,C", "D,D", "E,H", "I,I", "J,J" },
                list.Count);

            startRow += 3;
            for (int i = 0; i < list.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = list[i].Id;
                sheet.Cells[startRow + i, "C"] = list[i].Type;
                sheet.Cells[startRow + i, "D"] = list[i].Severity;
                sheet.Cells[startRow + i, "E"] = list[i].Title;
                sheet.Cells[startRow + i, "I"] = Utility.GetPersonName(list[i].AssignedTo);
                sheet.Cells[startRow + i, "J"] = Utility.GetPersonName(list[i].DiscoveryUser);
            }
            return nextRow;
        }
        private int BuildNoneTable(int startRow, List<BugEntity> list)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代处理的不是错误/不予处理Bug分析", "说明：", "B", "L",
                new List<string>() { "BugID", "关闭原因", "问题类别", "严重级别", "Bug标题", "不是错误/不予处理分析" },
                new List<string>() { "B,B", "C,C", "D,D", "E,E", "F,H", "I,L" },
                list.Count);

            startRow += 3;
            for (int i = 0; i < list.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = list[i].Id;
                sheet.Cells[startRow + i, "C"] = "";
                sheet.Cells[startRow + i, "D"] = list[i].Type;
                sheet.Cells[startRow + i, "E"] = list[i].Severity;
                sheet.Cells[startRow + i, "F"] = list[i].Title;
                sheet.Cells[startRow + i, "I"] = "";
            }

            return nextRow;
        }
    }
}
