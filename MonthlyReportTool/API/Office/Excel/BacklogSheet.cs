using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class BacklogSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public BacklogSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            BuildTitle();
            BuildSubTitle();
            BuildDescription();

            BuildSummaryTable();
            int startRow = BuildTable();
            startRow = BuildDelayedTable(startRow);
            BuildAbandonTable(startRow);
        }
        private void BuildTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "O"]];
            Utility.AddNativieResource(range);
            range.ColumnWidth = 8;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "Backlog统计分析";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
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
        private void BuildSummaryTable()
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
                { "已完成数", "", "", "已完成数：【已发布】及【已完成】状态的Backlog数\r\n占比：已完成数/本迭代计划总数"},
                { "进行中数", "", "", "进行中数：【测试通过】、【测试接收】、【开发完成】、【进行中】、【提交确认】状态的Backlog数\r\n占比：进行中数/本迭代计划总数"},
                { "未启动数", "", "", "未启动数：【已批准】、【提交评审】、【已承诺】、【新建】状态的Backlog数\r\n占比：未启动数/本迭代计划总数" },
                { "拖期数", "", "", "拖期数：进行中数+未启动数\r\n占比：拖期数/本迭代计划总数"},
                { "本迭代计划总数", "", "", "本迭代规划的所有backlog（包括上迭代拖期的）个数"},
                { "提交数", "", "", "提交数：【已发布】、【提交测试】、【测试接收】状态的Backlog数\r\n占比：提交数/应提交数" },
                { "未测试数", "", "", "未测试数：【进行中】、【开发完成】及其他状态的Backlog数\r\n占比：未测试数/应提交数"},
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

            sheet.Cells[7, "E"] = "=IF(D7<>0,D7/D11,\"\")";
            sheet.Cells[8, "E"] = "=IF(D8<>0,D8/D11,\"\")";
            sheet.Cells[9, "E"] = "=IF(D9<>0,D9/D11,\"\")";
            sheet.Cells[10, "E"] = "=IF(D10<>0,D10/D11,\"\")";
            sheet.Cells[11, "E"] = "'--";
            sheet.Cells[12, "E"] = "=IF(D12<>0,D12/D14,\"\")";
            sheet.Cells[13, "E"] = "=IF(D13<>0,D13/D14,\"\")";
            sheet.Cells[14, "E"] = "'--";
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

        private int BuildTable()
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[16, "B"], sheet.Cells[16, "M"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[16, "B"] = "本迭代backlog列表";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[17, "B"], sheet.Cells[17, "M"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[17, "B"] = "说明：按关键应用、模块排序；非研发类的为无";
            var tmpdesccharc = tableDescriptionRange.Characters[4, 10];

            var tmpdescfont = tmpdesccharc.Font;
            Utility.AddNativieResource(tmpdesccharc);
            Utility.AddNativieResource(tmpdescfont);
            tmpdescfont.Color = System.Drawing.Color.Red.ToArgb();

            int featuresCount = 20;
            string[] cols = new string[] { "ID", "关键应用", "模块", "backlog名称", "类别", "负责人", "状态" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","C"),
                Tuple.Create<string,string>("D","E"),
                Tuple.Create<string,string>("F","J"),
                Tuple.Create<string,string>("K","K"),
                Tuple.Create<string,string>("L","L"),
                Tuple.Create<string,string>("M","M"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[18, colsname[i].Item1], sheet.Cells[18, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[18, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[18, "B"], sheet.Cells[18, "M"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Name = "微软雅黑";
            tableFont.Size = 11;
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[19 + i, colsname[j].Item1], sheet.Cells[19 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[19 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 19 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return 18 + featuresCount + 2;
        }

        private int BuildDelayedTable(int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "M"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "拖期backlog分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "M"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[startRow + 1, "B"] = "说明：分析每个拖期Backlog的原因、主要责任人、以及拖期改进措施、改进措施责任人";

            int featuresCount = 10;
            string[] cols = new string[] { "ID", "backlog名称", "拖期责任人", "拖期原因", "拖期改进措施", "措施负责人" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","E"),
                Tuple.Create<string,string>("F","F"),
                Tuple.Create<string,string>("G","I"),
                Tuple.Create<string,string>("J","L"),
                Tuple.Create<string,string>("M","M"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 2, colsname[i].Item1], sheet.Cells[startRow + 2, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[startRow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2, "M"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Name = "微软雅黑";
            tableFont.Size = 11;
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 3 + i, colsname[j].Item1], sheet.Cells[startRow + 3 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[startRow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", startRow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return startRow + 3 + featuresCount + 2;
        }

        private int BuildAbandonTable(int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "M"]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "移除/中止backlog分析";
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "M"]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[startRow + 1, "B"] = "说明：分析每个移除/中止Backlog的处理原因";

            int featuresCount = 3;
            string[] cols = new string[] { "ID", "backlog名称", "移除/中止原因分析", "负责人" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","F"),
                Tuple.Create<string,string>("G","L"),
                Tuple.Create<string,string>("M","M"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 2, colsname[i].Item1], sheet.Cells[startRow + 2, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[startRow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2, "M"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            Utility.AddNativieResource(tableFont);
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 3 + i, colsname[j].Item1], sheet.Cells[startRow + 3 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[startRow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", startRow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return startRow + 3 + featuresCount + 2;
        }
    }
}
