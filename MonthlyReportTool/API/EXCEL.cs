using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;

namespace MonthlyReportTool.API
{
    public class EXCEL
    {
        private static List<object> nativeResources = new List<object>();
        public static void BuildIterationReports()
        {
            ExcelInterop.Application excel = new ExcelInterop.Application();

            excel.DisplayAlerts = false;
            ExcelInterop.Workbook workbook = excel.Workbooks.Add();
            nativeResources.Add(workbook);
            ExcelInterop.Sheets sheets = workbook.Worksheets;
            nativeResources.Add(sheets);

            var sheetHome = (ExcelInterop.Worksheet)sheets.Add();
            nativeResources.Add(sheetHome);
            sheetHome.Name = "首页及说明";
            BuildHomeSheet(sheetHome);

            var sheetContent = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetHome);
            nativeResources.Add(sheetContent);
            sheetContent.Name = "目录";
            BuildContentSheet(sheetContent);

            var sheetOverview = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetContent);
            nativeResources.Add(sheetOverview);
            sheetOverview.Name = "项目整体说明";
            BuildOverviewSheet(sheetOverview);

            var sheetFeature = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetOverview);
            nativeResources.Add(sheetFeature);
            sheetFeature.Name = "产品特性统计";
            BuildFeatureSheet(sheetFeature);

            var sheetBacklog = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetFeature);
            nativeResources.Add(sheetBacklog);
            sheetBacklog.Name = "Backlog统计";
            BuildBacklogSheet(sheetBacklog);

            var sheetWorkload = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetBacklog);
            nativeResources.Add(sheetWorkload);
            sheetWorkload.Name = "工作量统计";
            BuildWorkloadSheet(sheetWorkload);

            var sheetCommitment = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetWorkload);
            nativeResources.Add(sheetCommitment);
            sheetCommitment.Name = "提交单分析";
            BuildCommitmentSheet(sheetCommitment);

            var sheetCodeReview = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetCommitment);
            nativeResources.Add(sheetCodeReview);
            sheetCodeReview.Name = "代码审查分析";
            BuildCodeReviewSheet(sheetCodeReview);

            var sheetBugAnalysis = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetCodeReview);
            nativeResources.Add(sheetBugAnalysis);
            sheetBugAnalysis.Name = "Bug统计分析";
            BuildBugAnalysisSheet(sheetBugAnalysis);

            var sheetSuggestion = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetBugAnalysis);
            nativeResources.Add(sheetSuggestion);
            sheetSuggestion.Name = "改进建议";
            BuildSuggestionSheet(sheetSuggestion);

            var sheetPeoplePerformance = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetSuggestion);
            nativeResources.Add(sheetPeoplePerformance);
            sheetPeoplePerformance.Name = "人员考评结果";
            BuildPeoplePerformanceSheet(sheetPeoplePerformance);

            sheets.Select();//选择所有的sheet

            var window = excel.ActiveWindow;
            nativeResources.Add(window);
            window.DisplayGridlines = false;//都不显示表格线
            workbook.SaveAs("c:\\irt\\1.xlsx");
            workbook.Close();

            foreach (object com in nativeResources)
            {
                TFS.Utils.ReleaseComObject(com);
            }

            excel.Quit();
        }

        private static void BuildHomeSheet(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range allrange = sheet.Range[sheet.Cells[1, "C"], sheet.Cells[40, "J"]];
            nativeResources.Add(allrange);
            allrange.ColumnWidth = 12;
            allrange.RowHeight = 15;

            #region 1st paragraph
            sheet.Cells[5, "D"] = "***迭代总结";
            ExcelInterop.Range range = sheet.Range[sheet.Cells[5, "D"], sheet.Cells[6, "H"]];
            nativeResources.Add(range);
            range.Merge();
            var font = range.Font;
            font.Size = 20;
            font.Name = "微软雅黑";
            font.Bold = true;
            nativeResources.Add(font);
            #endregion 1st paragraph

            #region 2nd paragraph
            string text = "\r\n模板说明：\r\n" +
                            "      用途：用于各团队项目编制迭代总结的参考。\r\n" +
                            "      标题：团队项目全称 + SprintXX + 总结。例如：公共技术Sprint34总结。\r\n" +
                            "文档命名：团队项目简称 + SprintXX + 总结 +（报告期间）。例如：TTP Sprint34总结（20171009_20171028）\r\n" +
                            "      页眉：和文档标题一致。例如：公共技术Sprint34总结。\r\n" +
                            "      注解：正文中倾斜字体部分需要替换成实际内容，且修改为非倾斜字体。\r\n";
            sheet.Cells[10, "C"] = text;

            ExcelInterop.Range range2 = sheet.Range[sheet.Cells[10, "C"], sheet.Cells[17, "J"]];
            nativeResources.Add(range2);
            range2.Merge();
            range2.UseStandardHeight = true;
            range2.WrapText = true;
            range2.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            var font2 = range2.Font;
            nativeResources.Add(font2);
            font2.Size = 11;
            font2.Name = "微软雅黑";

            List<Tuple<int, int>> characters = new List<Tuple<int, int>>(){
                Tuple.Create<int, int>(3, 4),
                Tuple.Create<int, int>(16, 3),
                Tuple.Create<int, int>(44, 3),
                Tuple.Create<int, int>(90,5),
                Tuple.Create<int, int>(170,3),
                Tuple.Create<int, int>(207,3),
            };

            foreach (var charc in characters)
            {
                var tmpcharc = range2.Characters[charc.Item1, charc.Item2];
                var tmpfont = tmpcharc.Font;
                tmpfont.Bold = true;

                nativeResources.Add(tmpfont);
                nativeResources.Add(tmpcharc);
            }

            var border2 = range2.Borders;
            nativeResources.Add(border2);
            border2.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 2nd paragraph

            #region 3rd paragraph
            text = "表格内容使用说明：1、表格中需要公式计算的地方已添加公式，填入统计数据后，派生数据会自动生成，底色为绿色的不要随意改动；";
            sheet.Cells[19, "C"] = text;

            ExcelInterop.Range range3 = sheet.Range[sheet.Cells[19, "C"], sheet.Cells[23, "J"]];
            nativeResources.Add(range3);
            range3.Merge();
            range3.UseStandardHeight = true;
            range3.WrapText = true;
            range3.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            var font3 = range3.Font;
            nativeResources.Add(font3);
            font3.Size = 11;
            font3.Name = "微软雅黑";

            var tmpcharc3 = range3.Characters[1, 9];
            var tmpfont3 = tmpcharc3.Font;
            tmpfont3.Bold = true;

            nativeResources.Add(tmpfont3);
            nativeResources.Add(tmpcharc3);


            var border3 = range3.Borders;
            nativeResources.Add(border3);
            border3.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 3rd paragraph

            #region 4th paragraph
            text = "项目模板定义：1、此模板为组织级模板，各团队项目根据人员投入及团队项目情况定义自己的模板，团队项目模板生成后，每次直接填写数据即可，不必每次调整模板格式；";
            sheet.Cells[25, "C"] = text;

            ExcelInterop.Range range4 = sheet.Range[sheet.Cells[25, "C"], sheet.Cells[28, "J"]];
            nativeResources.Add(range4);
            range4.Merge();
            range4.UseStandardHeight = true;
            range4.WrapText = true;
            range4.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            var font4 = range4.Font;
            nativeResources.Add(font4);
            font4.Size = 11;
            font4.Name = "微软雅黑";

            var tmpcharc4 = range4.Characters[1, 9];
            var tmpfont4 = tmpcharc4.Font;
            tmpfont4.Bold = true;

            nativeResources.Add(tmpfont4);
            nativeResources.Add(tmpcharc4);


            var border4 = range4.Borders;
            nativeResources.Add(border4);
            border4.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 3rd paragraph
        }

        private static void BuildContentSheet(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range allrange = sheet.Range[sheet.Cells[5, "D"], sheet.Cells[40, "J"]];
            nativeResources.Add(allrange);
            allrange.ColumnWidth = 12;
            allrange.RowHeight = 15;
            allrange.Merge();
            allrange.UseStandardHeight = true;
            allrange.WrapText = true;
            allrange.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;

            sheet.Cells[5, "D"] = "这个目录没个鸟用。";
        }

        private static void BuildOverviewSheet(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "J"]];
            nativeResources.Add(titleRange);
            titleRange.ColumnWidth = 10;
            titleRange.RowHeight = 40;
            titleRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            titleRange.Merge();
            sheet.Cells[2, "B"] = "项目整体说明";
            var titleFont = titleRange.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            nativeResources.Add(colA);
            colA.ColumnWidth = 2;

            ExcelInterop.Range title2Range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "J"]];
            nativeResources.Add(title2Range);
            title2Range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignLeft;
            title2Range.Merge();

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[12, "B"]];
            nativeResources.Add(tableRange);
            var tableBorder = tableRange.Borders;
            nativeResources.Add(tableBorder);
            tableBorder.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            for (int row = 5; row <= 12; row++)
            {
                ExcelInterop.Range table2Range = sheet.Range[sheet.Cells[row, "C"], sheet.Cells[row, "J"]];
                nativeResources.Add(table2Range);
                var table2Border = table2Range.Borders;
                nativeResources.Add(table2Border);
                table2Border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                table2Range.Merge();
            }

            ExcelInterop.Range leftRange = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[12, "B"]];
            nativeResources.Add(leftRange);
            leftRange.ColumnWidth = 30.67;
            leftRange.RowHeight = 20;
            var leftFont = leftRange.Font;
            nativeResources.Add(leftFont);
            leftFont.Bold = true;
            leftFont.Name = "微软雅黑";
            leftFont.Size = 11;

            sheet.Cells[4, "B"] = "迭代期间及人员情况综述";
            sheet.Cells[5, "B"] = "Sprint期间";
            sheet.Cells[6, "B"] = "项目负责人";
            sheet.Cells[7, "B"] = "开发负责人";
            sheet.Cells[8, "B"] = "开发人员";
            sheet.Cells[9, "B"] = "需求人员";
            sheet.Cells[10, "B"] = "UI人员";
            sheet.Cells[11, "B"] = "测试负责人";
            sheet.Cells[12, "B"] = "测试人员";


        }

        private static void BuildFeatureSheet(ExcelInterop.Worksheet sheet)
        {
            #region Some Titles
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "O"]];
            nativeResources.Add(titleRange);
            titleRange.ColumnWidth = 8;
            titleRange.RowHeight = 40;
            titleRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            titleRange.Merge();
            sheet.Cells[2, "B"] = "产品特性统计分析";
            var titleFont = titleRange.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 20;

            ExcelInterop.Range title2Range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            nativeResources.Add(title2Range);
            //title2Range.RowHeight = 40;
            title2Range.Merge();
            sheet.Cells[4, "B"] = "本迭代产品特性完成情况统计";
            var title2Font = title2Range.Font;
            nativeResources.Add(title2Font);
            title2Font.Bold = true;
            title2Font.Name = "微软雅黑";
            title2Font.Size = 12;

            sheet.Cells[5, "B"] = "说明：统计依据为：完成本月目标状态的本月目标日期落在本迭代期间内的产品特性";

            ExcelInterop.Range title3Range = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "O"]];
            nativeResources.Add(title3Range);
            title3Range.Merge();

            var title3Font = title3Range.Font;
            nativeResources.Add(title3Font);
            title3Font.Name = "微软雅黑";
            title3Font.Size = 11;

            var tmpcharc3 = title3Range.Characters[1, 3];

            var tmpfont3 = tmpcharc3.Font;
            nativeResources.Add(tmpcharc3);
            nativeResources.Add(tmpfont3);
            tmpfont3.Bold = true;
            #endregion Some Titles

            #region Table 1
            string[,] cols = new string[,]
            {
                { "分类", "已完成数", "拖期数", "按计划完成数", "本迭代计划总数" },
                { "个数", "", "", "", "" },
                { "占比", "", "", "", "" },
                { "说明", "已完成数：已完成本月目标的产品特性个数\r\n占比：已完成数/本迭代计划总数", "拖期数：未完成本迭代目标的产品特性个数\r\n占比：拖期数/本迭代计划总数", "按计划已完成数：按本月目标日期完成的产品特性个数\r\n占比：按计划完成数/本迭代计划总数", "本迭代时间范围内所有迭代产品特性总数本迭代计划总数=已完成数+拖期数" },
            };
            for (int row = 6; row <= 10; row++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row, "B"], sheet.Cells[row, "D"]];
                nativeResources.Add(colRange);
                colRange.RowHeight = 40;
                colRange.Merge();
                sheet.Cells[row, "B"] = cols[0, row - 6];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                ExcelInterop.Range col2Range = sheet.Cells[row, "E"] as ExcelInterop.Range;
                nativeResources.Add(col2Range);
                col2Range.RowHeight = 40;
                col2Range.Merge();
                sheet.Cells[row, "E"] = cols[1, row - 6];

                var border2 = col2Range.Borders;
                nativeResources.Add(border2);
                border2.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                ExcelInterop.Range col3Range = sheet.Cells[row, "F"] as ExcelInterop.Range;
                nativeResources.Add(col3Range);
                col3Range.RowHeight = 40;
                col3Range.Merge();
                sheet.Cells[row, "F"] = cols[2, row - 6];

                var border3 = col3Range.Borders;
                nativeResources.Add(border3);
                border3.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                ExcelInterop.Range col4Range = sheet.Range[sheet.Cells[row, "G"], sheet.Cells[row, "O"]];
                nativeResources.Add(col4Range);
                col4Range.RowHeight = 40;
                col4Range.Merge();
                sheet.Cells[row, "G"] = cols[3, row - 6];

                var border4 = col4Range.Borders;
                nativeResources.Add(border4);
                border4.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range col5Range = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "O"]];
            nativeResources.Add(col5Range);
            col5Range.RowHeight = 20;
            col5Range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = col5Range.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var col5Font = col5Range.Font;
            nativeResources.Add(col5Font);
            col5Font.Name = "微软雅黑";
            col5Font.Size = 11;
            col5Font.Bold = true;

            sheet.Cells[7, "F"] = "=IF(E7<>0,E7/E10,\"\")";
            sheet.Cells[8, "F"] = "=IF(E8<>0,E8/E10,\"\")";
            sheet.Cells[9, "F"] = "=IF(E9<>0,E9/E10,\"\")";
            sheet.Cells[10, "E"] = "=SUM(E7: E8)";
            sheet.Cells[10, "F"] = "'--";
            #endregion Table 1

            #region Table 2
            ExcelInterop.Range table2TitleRange = sheet.Range[sheet.Cells[12, "B"], sheet.Cells[12, "O"]];
            nativeResources.Add(table2TitleRange);
            //title2Range.RowHeight = 40;
            table2TitleRange.Merge();
            sheet.Cells[12, "B"] = "本迭代产品特性列表";
            var table2TitleFont = table2TitleRange.Font;
            nativeResources.Add(table2TitleFont);
            table2TitleFont.Bold = true;
            table2TitleFont.Name = "微软雅黑";
            table2TitleFont.Size = 12;

            ExcelInterop.Range table2Title2Range = sheet.Range[sheet.Cells[13, "B"], sheet.Cells[13, "O"]];
            nativeResources.Add(table2Title2Range);
            //title2Range.RowHeight = 40;
            table2Title2Range.Merge();
            table2Title2Range.RowHeight = 60;
            sheet.Cells[13, "B"] = "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      如果本迭代实际的产品特性数多于模板预制的行数，请自行插入行，然后用格式刷刷新增的行的格式\r\n      按关键应用、模块排序；非研发类的为无";
            var tmpcharc23 = table2Title2Range.Characters[1, 3];

            var tmpfont23 = tmpcharc23.Font;
            nativeResources.Add(tmpcharc23);
            nativeResources.Add(tmpfont23);
            tmpfont23.Bold = true;

            int featuresCount = 20;
            string[] cols2 = new string[] { "ID", "关键应用", "模块", "产品特性名称", "目标状态", "目标日期", "负责人", "当前状态" };
            List<Tuple<string, string>> range = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","D"),
                Tuple.Create<string,string>("E","F"),
                Tuple.Create<string,string>("G","I"),
                Tuple.Create<string,string>("J","K"),
                Tuple.Create<string,string>("L","M"),
                Tuple.Create<string,string>("N","N"),
                Tuple.Create<string,string>("O","O"),
            };

            for (int i = 0; i < cols2.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[14, range[i].Item1], sheet.Cells[14, range[i].Item2]];
                nativeResources.Add(colRange);
                colRange.RowHeight = 40;
                colRange.Merge();
                sheet.Cells[14, range[i].Item1] = cols2[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range table2Range = sheet.Range[sheet.Cells[14, "B"], sheet.Cells[14, "O"]];
            nativeResources.Add(table2Range);
            table2Range.RowHeight = 40;
            table2Range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior2 = table2Range.Interior;
            nativeResources.Add(interior2);
            interior2.Color = System.Drawing.Color.DarkGray.ToArgb();

            var table2Font = table2Range.Font;
            nativeResources.Add(table2Font);
            table2Font.Name = "微软雅黑";
            table2Font.Size = 11;
            table2Font.Bold = true;

            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols2.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[15+i, range[j].Item1], sheet.Cells[15 + i, range[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[15 + i, range[j].Item1] = String.Format("数据行:{0}，列{1}", 15 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }
            #endregion Table 2
        }

        private static void BuildBacklogSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildWorkloadSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildCommitmentSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildCodeReviewSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildBugAnalysisSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildSuggestionSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildPeoplePerformanceSheet(ExcelInterop.Worksheet sheet)
        {

        }

    }
}
