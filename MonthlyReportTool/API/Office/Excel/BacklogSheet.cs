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
    public class BacklogSheet : ExcelSheetBase, IExcelSheet
    {
        private List<List<BacklogEntity>> backlogList;
        private ExcelInterop.Worksheet sheet;
        public BacklogSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            this.backlogList = Backlog.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));
            BuildTitle();
            BuildSubTitle();
            BuildDescription();

            int startRow = BuildSummaryTable();
            startRow = BuildDelayedTable(startRow);
            startRow = BuildAbandonTable(startRow);
            BuildTable(startRow);
        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "O", "Backlog统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "本迭代所有计划backlog完成情况统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 12;

        }
        private void BuildDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "O"]];
            Utility.AddNativieResource(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Name = "微软雅黑";
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;
        }
        private int BuildSummaryTable()
        {
            var rb = sheet.Cells[1, "B"] as ExcelInterop.Range;
            rb.ColumnWidth = 10;
            Utility.AddNativieResource(rb);
            var rc = sheet.Cells[1, "C"] as ExcelInterop.Range;
            rc.ColumnWidth = 20;
            Utility.AddNativieResource(rc);

            string[,] cols = new string[,]
                        {
                { "分类", "个数", "占比", "说明"},
                { "已完成数", "", "", "已完成数：【已发布】及【已完成】状态的、或者根据验收标准已完成的Backlog总数\r\n占比：已完成数 / 本迭代计划总数"},
                { "进行中数", "", "", "进行中数：本迭代计划总数-已完成数-未启动数-中止/移除数\r\n占比：进行中数 / 本迭代计划总数"},
                { "未启动数", "", "", "未启动数：【已承诺】、【新建】、【已批准】、【提交评审】状态的Backlog总数\r\n占比：未启动数 / 本迭代计划总数" },
                { "拖期数", "", "", "拖期数：本迭代计划总数-已完成数-中止数/移除数\r\n占比：拖期数 / 本迭代计划总数"	},
                { "中止/移除数", "", "", "中止/移除数：本迭代已中止或已移除的Backlog总数\r\n占比：移除数 / 本迭代计划总数"},


                { "本迭代计划总数", "", "", "本迭代规划的所有backlog总数（包括当前已经被移除、中止的）"},
                { "提交测试数", "", "", "提交测试数：【已完成】、【已发布】、【提交测试】、【测试接收】、【测试通过】状态的Backlog总数\r\n占比：提交测试数 / 应提交数"	 },
                { "测试通过数", "", "", "测试通过数：【已发布】、【测试通过】、【已完成】状态的Backlog总数\r\n占比：测试通过数 / 应提交数"},
                { "未提交或未测试通过数", "", "", "未提交或未测试通过数：应提交数-提交测试数\r\n占比：未提交或未测试通过数 / 应提交数"	},
                { "应提交数", "", "", "应提交数：本迭代Backlog类别是【开发】、完成标准为【测试通过】及【发布上线】的Backlog总数\r\n未测试及提交数都是以这两个条件为基本过滤"},
                        };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","C"),
                Tuple.Create<string,string>("D","D"),
                Tuple.Create<string,string>("E","E"),
                Tuple.Create<string,string>("F","K"),
            };
            for (int row = 0; row < cols.GetLength(0); row++)
            {
                for (int col = 0; col < colsname.Count; col++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[6 + row, colsname[col].Item1], sheet.Cells[6 + row, colsname[col].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 40;
                    colRange.Merge();
                    sheet.Cells[6 + row, colsname[col].Item1] = cols[row, col];

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                    if (col == 3)
                    {
                        colRange.ColumnWidth = 12;
                    }
                }
            }

            BuildSummaryTableTitle();

            sheet.Cells[7, "E"] = "=IF(D7<>0,D7/D12,\"\")";
            sheet.Cells[8, "D"] = "=D12-D7-D9-D11";
            sheet.Cells[8, "E"] = "=IF(D8<>0,D8/D12,\"\")";
            sheet.Cells[9, "E"] = "=IF(D9<>0,D9/D12,\"\")";
            sheet.Cells[10, "D"] = "=D12-D7-D11";
            sheet.Cells[10, "E"] = "=IF(D10<>0,D10/D12,\"\")";
            sheet.Cells[11, "E"] = "=IF(D11<>0,D11/D12,\"\")";

            sheet.Cells[12, "E"] = "'--";
            sheet.Cells[13, "E"] = "=IF(D13<>0,D13/D16,\"\")";
            sheet.Cells[14, "E"] = "=IF(D14<>0,D14/D16,\"\")";
            sheet.Cells[15, "D"] = "=D16-D13";
            sheet.Cells[15, "E"] = "=IF(D15<>0,D15/D16,\"\")";
            sheet.Cells[16, "E"] = "'--";

            Utility.SetupSheetPercentFormat(sheet, 7, "E", 16, "E");

            sheet.Cells[7, "D"] = this.backlogList[0].Count;
            sheet.Cells[9, "D"] = this.backlogList[1].Count;
            sheet.Cells[11, "D"] = this.backlogList[2].Count;
            sheet.Cells[12, "D"] = this.backlogList[3].Count;
            sheet.Cells[13, "D"] = this.backlogList[4].Count;
            sheet.Cells[14, "D"] = this.backlogList[5].Count;
            sheet.Cells[16, "D"] = this.backlogList[6].Count;
            

            return 17;
        }

        private void BuildSummaryTableTitle()
        {
            ExcelInterop.Range colRange = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "K"]];
            Utility.AddNativieResource(colRange);
            colRange.RowHeight = 20;
            colRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = colRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var colFont = colRange.Font;
            Utility.AddNativieResource(colFont);
            colFont.Bold = true;
        }

        

        private int BuildDelayedTable(int startRow)
        {
            //拖期数：本迭代计划总数-已完成数-中止数-移除数
            #region 有了查询了，不用自己算了
            //var all = this.backlogList[3];
            //var done = this.backlogList[0];
            //var abandon = this.backlogList[2];

            //for (int i = 0; i < all.Count; i++)
            //{
            //    var matched = done.Where(backlog => backlog.Id == all[i].Id);
            //    if (matched.Count() > 0)
            //    {                    
            //        done.Remove(matched.First());
            //        goto HELLOWORLD;

            //    }

            //    matched = abandon.Where(backlog => backlog.Id == all[i].Id);
            //    if (matched.Count() > 0)
            //    {                    
            //        abandon.Remove(matched.First());
            //        goto HELLOWORLD;
            //    }

            //    continue;

            //    HELLOWORLD:
            //    {
            //        all.RemoveAt(i);
            //        i = 0;
            //    }
            //}
            #endregion 有了查询了，不用自己算了

            var all = this.backlogList[7];
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "拖期backlog分析", "说明：分析每个拖期Backlog的原因、主要责任人、以及拖期改进措施、改进措施责任人", "B", "M",
                new List<string>() { "ID", "关键应用", "模块", "backlog名称", "类别", "负责人", "验收标准", "状态", "拖期责任人", "拖期原因", "拖期改进措施", "措施负责人" },
                new List<string>() { "B,B", "C,C", "D,E", "F,J", "K,K", "L,L", "M,N", "O,O", "P,P", "Q,T", "U,X", "Y,Y" },
                all.Count);

            

            for (int i = 0; i < all.Count; i++)
            {
                sheet.Cells[startRow + 3 + i, "B"] = all[i].Id;
                sheet.Cells[startRow + 3 + i, "C"] = all[i].KeyApplication;
                sheet.Cells[startRow + 3 + i, "D"] = all[i].ModulesName;
                sheet.Cells[startRow + 3 + i, "F"] = all[i].Title;
                sheet.Cells[startRow + 3 + i, "K"] = all[i].Category;
                sheet.Cells[startRow + 3 + i, "L"] = all[i].AssignedTo;
                sheet.Cells[startRow + 3 + i, "M"] = all[i].AcceptanceMeasure;
                sheet.Cells[startRow + 3 + i, "O"] = all[i].State;
                sheet.Cells[startRow + 3 + i, "P"] = "";
                sheet.Cells[startRow + 3 + i, "Q"] = "";
                sheet.Cells[startRow + 3 + i, "U"] = "";
                sheet.Cells[startRow + 3 + i, "Y"] = "";
            }
            return nextRow;
        }

        private int BuildAbandonTable(int startRow)
        {
            var all = this.backlogList[2];
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "移除/中止backlog分析", "说明：分析每个移除/中止Backlog的处理原因", "B", "M",
                new List<string>() { "ID", "关键应用", "模块", "backlog名称", "类别", "负责人", "验收标准","状态", "移除/中止原因分析" },
                new List<string>() { "B,B", "C,C", "D,E", "F,J","K,K","L,L","M,N","O,O","P,T" },
                all.Count);

            for (int i = 0; i < all.Count; i++)
            {
                sheet.Cells[startRow + 3 + i, "B"] = all[i].Id;
                sheet.Cells[startRow + 3 + i, "C"] = all[i].KeyApplication;
                sheet.Cells[startRow + 3 + i, "D"] = all[i].ModulesName;
                sheet.Cells[startRow + 3 + i, "F"] = all[i].Title;
                sheet.Cells[startRow + 3 + i, "K"] = all[i].Category;
                sheet.Cells[startRow + 3 + i, "L"] = all[i].AssignedTo;
                sheet.Cells[startRow + 3 + i, "M"] = all[i].AcceptanceMeasure;
                sheet.Cells[startRow + 3 + i, "O"] = all[i].State;
                sheet.Cells[startRow + 3 + i, "P"] = "";
            }
            return nextRow;
        }

        private int BuildTable(int startRow)
        {
            var all = this.backlogList[3];
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代backlog列表", "说明：按关键应用、模块排序；非研发类的为无", "B", "M",
                new List<string>() { "ID", "关键应用", "模块", "backlog名称", "类别", "负责人", "验收标准", "状态" },
                new List<string>() { "B,B", "C,C", "D,E", "F,J", "K,K", "L,L", "M,N","O,O" },
                all.Count);

            for (int i = 0; i < all.Count; i++)
            {
                sheet.Cells[startRow + 3 + i, "B"] = all[i].Id;
                sheet.Cells[startRow + 3 + i, "C"] = all[i].KeyApplication;
                sheet.Cells[startRow + 3 + i, "D"] = all[i].ModulesName;
                sheet.Cells[startRow + 3 + i, "F"] = all[i].Title;
                sheet.Cells[startRow + 3 + i, "K"] = all[i].Category;
                sheet.Cells[startRow + 3 + i, "L"] = all[i].AssignedTo;
                sheet.Cells[startRow + 3 + i, "M"] = all[i].AcceptanceMeasure;
                sheet.Cells[startRow + 3 + i, "O"] = all[i].State;
                sheet.Cells[startRow + 3 + i, "P"] = "";
            }

            return nextRow;
        }
    }
}
