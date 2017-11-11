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

        private static void SetSheetFont(ExcelInterop.Worksheet sheet)
        {
            var bigrange = sheet.Range[sheet.Cells[1, "A"], sheet.Cells[1000, "Z"]];
            nativeResources.Add(bigrange);
            var bigrangeFont = bigrange.Font;
            nativeResources.Add(bigrangeFont);

            bigrangeFont.Name = "微软雅黑";
            bigrangeFont.Size = 11;
        }
        private static void BuildHomeSheet(ExcelInterop.Worksheet sheet)
        {
            SetSheetFont(sheet);
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
            SetSheetFont(sheet);
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
            SetSheetFont(sheet);
            #region 标题
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
            #endregion 标题

            #region 标题2
            ExcelInterop.Range title2Range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "J"]];
            nativeResources.Add(title2Range);
            title2Range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignLeft;
            title2Range.Merge();

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[12, "B"]];
            nativeResources.Add(tableRange);
            var tableBorder = tableRange.Borders;
            nativeResources.Add(tableBorder);
            tableBorder.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 标题2

            #region 表格
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

            #endregion 表格
        }

        #region 产品特性
        private static void BuildFeatureSheet(ExcelInterop.Worksheet sheet)
        {
            SetSheetFont(sheet);
            BuildFeatureTitle(sheet);
            BuildFeatureSubTitle(sheet);
            BuildFeatureDescription(sheet);

            BuildFeatureSummaryTable(sheet);
            int featuresCount = BuildFeatureTable(sheet);
            featuresCount = BuildAnadonFeatureTable(sheet, featuresCount);
        }


        private static void BuildFeatureTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "O"]];
            nativeResources.Add(range);
            range.ColumnWidth = 8;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "产品特性统计分析";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            nativeResources.Add(colA);
            colA.ColumnWidth = 2;
        }
        private static void BuildFeatureSubTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            nativeResources.Add(range);
            range.Merge();
            sheet.Cells[4, "B"] = "本迭代产品特性完成情况统计";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 12;

        }
        private static void BuildFeatureDescription(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "O"]];
            nativeResources.Add(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明：统计依据为：完成本月目标状态的本月目标日期落在本迭代期间内的产品特性";

            var titleFont = titleRange.Font;
            nativeResources.Add(titleFont);
            titleFont.Name = "微软雅黑";
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            nativeResources.Add(tmpchar);
            nativeResources.Add(tmpfont);
            tmpfont.Bold = true;
        }
        private static void BuildFeatureSummaryTable(ExcelInterop.Worksheet sheet)
        {
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
        }

        private static int BuildFeatureTable(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[12, "B"], sheet.Cells[12, "O"]];
            nativeResources.Add(tableTitleRange);
            //title2Range.RowHeight = 40;
            tableTitleRange.Merge();
            sheet.Cells[12, "B"] = "本迭代产品特性列表";
            var tableTitleFont = tableTitleRange.Font;
            nativeResources.Add(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Name = "微软雅黑";
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[13, "B"], sheet.Cells[13, "O"]];
            nativeResources.Add(tableDescriptionRange);
            //title2Range.RowHeight = 40;
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 60;
            sheet.Cells[13, "B"] = "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      如果本迭代实际的产品特性数多于模板预制的行数，请自行插入行，然后用格式刷刷新增的行的格式\r\n      按关键应用、模块排序；非研发类的为无";
            var tmpdesccharc = tableDescriptionRange.Characters[1, 3];

            var tmpdescfont = tmpdesccharc.Font;
            nativeResources.Add(tmpdesccharc);
            nativeResources.Add(tmpdescfont);
            tmpdescfont.Bold = true;

            int featuresCount = 20;
            string[] cols = new string[] { "ID", "关键应用", "模块", "产品特性名称", "目标状态", "目标日期", "负责人", "当前状态" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","D"),
                Tuple.Create<string,string>("E","F"),
                Tuple.Create<string,string>("G","I"),
                Tuple.Create<string,string>("J","K"),
                Tuple.Create<string,string>("L","M"),
                Tuple.Create<string,string>("N","N"),
                Tuple.Create<string,string>("O","O"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[14, colsname[i].Item1], sheet.Cells[14, colsname[i].Item2]];
                nativeResources.Add(colRange);
                colRange.RowHeight = 40;
                colRange.Merge();
                sheet.Cells[14, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[14, "B"], sheet.Cells[14, "O"]];
            nativeResources.Add(tableRange);
            tableRange.RowHeight = 40;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            nativeResources.Add(tableFont);
            tableFont.Name = "微软雅黑";
            tableFont.Size = 11;
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[15 + i, colsname[j].Item1], sheet.Cells[15 + i, colsname[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[15 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 15 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return featuresCount;
        }
        private static int BuildAnadonFeatureTable(ExcelInterop.Worksheet sheet, int featuresCount)
        {
            int nextrow = 15 + featuresCount + 1;
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[nextrow, "B"], sheet.Cells[nextrow, "O"]];
            nativeResources.Add(tableTitleRange);
            //title2Range.RowHeight = 40;
            tableTitleRange.Merge();
            sheet.Cells[nextrow, "B"] = "本迭代移除/中止产品特性分析";
            var tableTitleFont = tableTitleRange.Font;
            nativeResources.Add(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Name = "微软雅黑";
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[nextrow + 1, "B"], sheet.Cells[nextrow + 1, "O"]];
            nativeResources.Add(tableDescriptionRange);
            //title2Range.RowHeight = 40;
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 60;
            sheet.Cells[nextrow + 1, "B"] = "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      如果本迭代实际的产品特性数多于模板预制的行数，请自行插入行，然后用格式刷刷新增的行的格式\r\n      按关键应用、模块排序；非研发类的为无";
            var tmpdesccharc = tableDescriptionRange.Characters[1, 3];

            var tmpdescfont = tmpdesccharc.Font;
            nativeResources.Add(tmpdesccharc);
            nativeResources.Add(tmpdescfont);
            tmpdescfont.Bold = true;

            featuresCount = 10;
            string[] cols = new string[] { "ID", "关键应用", "模块", "产品特性名称", "负责人", "移除/中止原因说明" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","B"),
                Tuple.Create<string,string>("C","D"),
                Tuple.Create<string,string>("E","F"),
                Tuple.Create<string,string>("G","I"),
                Tuple.Create<string,string>("J","K"),
                Tuple.Create<string,string>("L","O"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[nextrow + 2, colsname[i].Item1], sheet.Cells[nextrow + 2, colsname[i].Item2]];
                nativeResources.Add(colRange);
                colRange.RowHeight = 40;
                colRange.Merge();
                sheet.Cells[nextrow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[nextrow + 2, "B"], sheet.Cells[nextrow + 2, "O"]];
            nativeResources.Add(tableRange);
            tableRange.RowHeight = 40;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            nativeResources.Add(tableFont);
            tableFont.Name = "微软雅黑";
            tableFont.Size = 11;
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[nextrow + 3 + i, colsname[j].Item1], sheet.Cells[nextrow + 3 + i, colsname[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[nextrow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", nextrow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return featuresCount;
        }
        #endregion 产品特性

        #region backlog
        private static void BuildBacklogSheet(ExcelInterop.Worksheet sheet)
        {
            SetSheetFont(sheet);
            BuildBacklogTitle(sheet);
            BuildBacklogSubTitle(sheet);
            BuildBacklogDescription(sheet);

            BuildBacklogSummaryTable(sheet);
            int startRow = BuildBacklogTable(sheet);
            startRow = BuildDelayedBacklogTable(sheet, startRow);
            BuildAbandonBacklogTable(sheet, startRow);
        }
        private static void BuildBacklogTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "O"]];
            nativeResources.Add(range);
            range.ColumnWidth = 8;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "Backlog统计分析";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            nativeResources.Add(colA);
            colA.ColumnWidth = 2;
        }
        private static void BuildBacklogSubTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            nativeResources.Add(range);
            range.Merge();
            sheet.Cells[4, "B"] = "本迭代所有计划backlog完成情况统计";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 12;

        }
        private static void BuildBacklogDescription(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "O"]];
            nativeResources.Add(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明";

            var titleFont = titleRange.Font;
            nativeResources.Add(titleFont);
            titleFont.Name = "微软雅黑";
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            nativeResources.Add(tmpchar);
            nativeResources.Add(tmpfont);
            tmpfont.Bold = true;
        }
        private static void BuildBacklogSummaryTable(ExcelInterop.Worksheet sheet)
        {
            var rb = sheet.Cells[1, "B"] as ExcelInterop.Range;
            rb.ColumnWidth = 10;
            nativeResources.Add(rb);
            var rc = sheet.Cells[1, "C"] as ExcelInterop.Range;
            rc.ColumnWidth = 20;
            nativeResources.Add(rc);

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
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 40;
                    colRange.Merge();
                    sheet.Cells[6 + row, colsname[col].Item1] = cols[row, col];

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                    if (col == 3)
                    {
                        colRange.ColumnWidth = 12;
                    }
                }
            }

            BuildBacklogSummaryTableTitle(sheet);

            sheet.Cells[7, "E"] = "=IF(D7<>0,D7/D11,\"\")";
            sheet.Cells[8, "E"] = "=IF(D8<>0,D8/D11,\"\")";
            sheet.Cells[9, "E"] = "=IF(D9<>0,D9/D11,\"\")";
            sheet.Cells[10, "E"] = "=IF(D10<>0,D10/D11,\"\")";
            sheet.Cells[11, "E"] = "'--";
            sheet.Cells[12, "E"] = "=IF(D12<>0,D12/D14,\"\")";
            sheet.Cells[13, "E"] = "=IF(D13<>0,D13/D14,\"\")";
            sheet.Cells[14, "E"] = "'--";
        }

        private static void BuildBacklogSummaryTableTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range colRange = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "K"]];
            nativeResources.Add(colRange);
            colRange.RowHeight = 20;
            colRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = colRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var colFont = colRange.Font;
            nativeResources.Add(colFont);
            colFont.Bold = true;
        }

        private static int BuildBacklogTable(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[16, "B"], sheet.Cells[16, "M"]];
            nativeResources.Add(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[16, "B"] = "本迭代backlog列表";
            var tableTitleFont = tableTitleRange.Font;
            nativeResources.Add(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[17, "B"], sheet.Cells[17, "M"]];
            nativeResources.Add(tableDescriptionRange);
            tableDescriptionRange.Merge();
            tableDescriptionRange.RowHeight = 20;
            sheet.Cells[17, "B"] = "说明：按关键应用、模块排序；非研发类的为无";
            var tmpdesccharc = tableDescriptionRange.Characters[4, 10];

            var tmpdescfont = tmpdesccharc.Font;
            nativeResources.Add(tmpdesccharc);
            nativeResources.Add(tmpdescfont);
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
                nativeResources.Add(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[18, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[18, "B"], sheet.Cells[18, "M"]];
            nativeResources.Add(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            nativeResources.Add(tableFont);
            tableFont.Name = "微软雅黑";
            tableFont.Size = 11;
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[19 + i, colsname[j].Item1], sheet.Cells[19 + i, colsname[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[19 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 19 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return 18 + featuresCount + 2;
        }

        private static int BuildDelayedBacklogTable(ExcelInterop.Worksheet sheet, int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "M"]];
            nativeResources.Add(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "拖期backlog分析";
            var tableTitleFont = tableTitleRange.Font;
            nativeResources.Add(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "M"]];
            nativeResources.Add(tableDescriptionRange);
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
                nativeResources.Add(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[startRow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2, "M"]];
            nativeResources.Add(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            nativeResources.Add(tableFont);
            tableFont.Name = "微软雅黑";
            tableFont.Size = 11;
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 3 + i, colsname[j].Item1], sheet.Cells[startRow + 3 + i, colsname[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[startRow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", startRow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return startRow + 3 + featuresCount + 2;
        }

        private static int BuildAbandonBacklogTable(ExcelInterop.Worksheet sheet, int startRow)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "M"]];
            nativeResources.Add(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[startRow, "B"] = "移除/中止backlog分析";
            var tableTitleFont = tableTitleRange.Font;
            nativeResources.Add(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[startRow + 1, "B"], sheet.Cells[startRow + 1, "M"]];
            nativeResources.Add(tableDescriptionRange);
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
                nativeResources.Add(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[startRow + 2, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2, "M"]];
            nativeResources.Add(tableRange);
            tableRange.RowHeight = 20;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var tableFont = tableRange.Font;
            nativeResources.Add(tableFont);
            tableFont.Bold = true;

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[startRow + 3 + i, colsname[j].Item1], sheet.Cells[startRow + 3 + i, colsname[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[startRow + 3 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", startRow + 3 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return startRow + 3 + featuresCount + 2;
        }
        #endregion backlog

        #region workload
        private static void BuildWorkloadSheet(ExcelInterop.Worksheet sheet)
        {
            SetSheetFont(sheet);
            BuildWorkloadTitle(sheet);
            BuildWorkloadSubTitle(sheet);
            BuildWorkloadDescription(sheet);
            BuildWorkloadSummaryTable(sheet);

            BuildDevelopmentWorkloadTitle(sheet);
            BuildDevelopmentWorkloadDescription(sheet);

            int startRow = BuildDevelopmentWorkloadTableTitle(sheet);

        }
        private static void BuildWorkloadTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[2, "B"], sheet.Cells[2, "AG"]];
            nativeResources.Add(range);
            range.ColumnWidth = 8;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[2, "B"] = "工作量统计分析";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            nativeResources.Add(colA);
            colA.ColumnWidth = 2;
        }
        private static void BuildWorkloadSubTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            nativeResources.Add(range);
            range.Merge();
            sheet.Cells[4, "B"] = "团队整体工作量分析";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 12;

        }
        private static void BuildWorkloadDescription(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "AG"]];
            nativeResources.Add(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明：容量投入总工时：迭代成员迭代规划容量×迭代天数之和计算";

            var titleFont = titleRange.Font;
            nativeResources.Add(titleFont);
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            nativeResources.Add(tmpchar);
            nativeResources.Add(tmpfont);
            tmpfont.Bold = true;
        }
        private static int BuildWorkloadSummaryTable(ExcelInterop.Worksheet sheet)
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
                nativeResources.Add(colRange);
                colRange.Merge();
                sheet.Cells[6, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                nativeResources.Add(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "W"]];
            nativeResources.Add(tableRange);
            tableRange.RowHeight = 50;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[7 + i, colsname[j].Item1], sheet.Cells[7 + i, colsname[j].Item2]];
                    nativeResources.Add(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[7 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 7 + i, j + 1);

                    var border = colRange.Borders;
                    nativeResources.Add(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return 7 + featuresCount + 2;
        }

        private static void BuildDevelopmentWorkloadTitle(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[9, "B"], sheet.Cells[9, "AG"]];
            nativeResources.Add(range);
            range.RowHeight = 20;
            range.Merge();
            sheet.Cells[9, "B"] = "开发工作量统计";
            var titleFont = range.Font;
            nativeResources.Add(titleFont);
            titleFont.Bold = true;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            nativeResources.Add(colA);
            colA.ColumnWidth = 2;
        }

        private static void BuildDevelopmentWorkloadDescription(ExcelInterop.Worksheet sheet)
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[10, "B"], sheet.Cells[10, "K"]];
            nativeResources.Add(titleRange);
            titleRange.RowHeight = 120;
            titleRange.Merge();
            sheet.Cells[10, "B"] = "说明：包括测试人员的工作量统计\r\n" +
           "        开发人员按照团队成员的研发工作量占比排序\r\n" +
           "        测试人员按照团队成员的Bug产出率排序\r\n" +
           "        各明细工作量是实际填写的工作日志某类的工作量，和工作日志中分类一致\r\n" +
           "        Bug数：对于开发人员是本迭代被测试出的Bug总数\r\n" +
           "                 对于测试人员是本迭代测试出的Bug总数";

            var titleFont = titleRange.Font;
            nativeResources.Add(titleFont);
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            nativeResources.Add(tmpchar);
            nativeResources.Add(tmpfont);
            tmpfont.Bold = true;

            var tmpchar2 = titleRange.Characters[30, 45];

            var tmpfont2 = tmpchar2.Font;
            nativeResources.Add(tmpchar2);
            nativeResources.Add(tmpfont2);
            tmpfont2.Color = System.Drawing.Color.Red.ToArgb();

            ExcelInterop.Range titleRange2 = sheet.Range[sheet.Cells[10, "L"], sheet.Cells[10, "W"]];
            titleRange2.Merge();
            nativeResources.Add(titleRange2);
            titleRange.Merge();
            sheet.Cells[10, "L"] = "      标准工作量：迭代天数*8\r\n" +
                                   "      实际投入工作量：不包含请假实际的工作量\r\n" +
                                   "      实际饱和度：实际投入工作量 / 标准工作量\r\n" +
                                   "      Bug产出率：bug数 / 实际投入工作量";
        }

        private static int BuildDevelopmentWorkloadTableTitle(ExcelInterop.Worksheet sheet)
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
Tuple.Create<int, string, int, string>(12,"R",12,"S"),Tuple.Create<int, string, int, string>(12,"T",12,"U"),

            };

            foreach (var tuple in allMergedCells)
            {
                ExcelInterop.Range range = sheet.Range[sheet.Cells[tuple.Item1, tuple.Item2], sheet.Cells[tuple.Item3, tuple.Item4]];
                nativeResources.Add(range);
                range.Merge();

                //var border = range.Borders;
                //nativeResources.Add(border);
                //border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }


            ExcelInterop.Range rangemini = sheet.Range[sheet.Cells[11, "B"], sheet.Cells[13, "AG"]];
            rangemini.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            rangemini.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            rangemini.WrapText = true;

            var bordermini = rangemini.Borders;
            nativeResources.Add(bordermini);
            bordermini.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            var interior = rangemini.Interior;
            nativeResources.Add(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

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
                sheet.Cells[14 + i, "B"] = String.Format("大荣荣{0:d3}", i + 1);
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
            nativeResources.Add(borderDevRange);
            borderDevRange.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            ExcelInterop.Range devRange2 = sheet.Range[sheet.Cells[14, "B"], sheet.Cells[14 + devWorkloadCount - 1, "B"]];
            nativeResources.Add(devRange2);
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
                sheet.Cells[14 + devWorkloadCount + 2 + i, "B"] = String.Format("大荣荣测试{0:d3}", i + 1);
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
            nativeResources.Add(borderTestRange);
            borderTestRange.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            sheet.Cells[14 + devWorkloadCount + 2 + testWorkloadCount, "B"] = "合计";

            int chartstart = 14 + devWorkloadCount + 2 + testWorkloadCount + 2;

            ExcelInterop.Range workloadChartRange = sheet.Range[sheet.Cells[chartstart, "B"], sheet.Cells[chartstart + 10, "AG"]];

            ExcelInterop.ChartObjects charts = sheet.ChartObjects(Type.Missing) as ExcelInterop.ChartObjects;
            nativeResources.Add(charts);

            ExcelInterop.ChartObject workloadChartObject = charts.Add(0, 0, workloadChartRange.Width, workloadChartRange.Height);
            nativeResources.Add(workloadChartObject);
            ExcelInterop.Chart workloadChart = workloadChartObject.Chart;//设置图表数据区域。
            nativeResources.Add(workloadChart);

            //=工作量统计!$B$14:$B$33,工作量统计!$F$14:$F$33
            ExcelInterop.Range datasource = sheet.get_Range(String.Format("B14:B{0},F14:F{0}", 14 + devWorkloadCount + 2 + testWorkloadCount - 1));//不是："B14:B25","F14:F25"
            //ExcelInterop.Range datasource = sheet.get_Range("B14:B25,F14:F25");//不是："B14:B25","F14:F25"
            nativeResources.Add(datasource);
            workloadChart.ChartWizard(datasource, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, "工作量饱和度", "人员", "工作量", Type.Missing);
            workloadChart.ApplyDataLabels();//图形上面显示具体的值
            //将图表移到数据区域之下。
            workloadChartObject.Left = Convert.ToDouble(workloadChartRange.Left);
            workloadChartObject.Top = Convert.ToDouble(workloadChartRange.Top) + 20;

            chartstart += 12;
            ExcelInterop.Range bugChartRange = sheet.Range[sheet.Cells[chartstart, "B"], sheet.Cells[chartstart + 10, "K"]];
            ExcelInterop.ChartObject bugChartObject = charts.Add(0, 0, bugChartRange.Width, bugChartRange.Height);
            nativeResources.Add(bugChartObject);
            ExcelInterop.Chart bugChart = bugChartObject.Chart;//设置图表数据区域。
            nativeResources.Add(bugChart);

            //=工作量统计!$B$14:$B$33,工作量统计!$F$14:$F$33
            ExcelInterop.Range datasource2 = sheet.get_Range(String.Format("B{0}:B{1},G{0}:G{1}", 14 + devWorkloadCount + 2 - 1, 14 + devWorkloadCount + 2 + testWorkloadCount - 1));//不是："B14:B25","F14:F25"
            //ExcelInterop.Range datasource = sheet.get_Range("B14:B25,F14:F25");//不是："B14:B25","F14:F25"
            nativeResources.Add(datasource2);
            bugChart.ChartWizard(datasource2, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, "测试人员BUG数", "人员", "BUG数", Type.Missing);
            bugChart.ApplyDataLabels();//图形上面显示具体的值
            //将图表移到数据区域之下。
            bugChartObject.Left = Convert.ToDouble(bugChartRange.Left);
            bugChartObject.Top = Convert.ToDouble(bugChartRange.Top) + 20;

            return 14 + devWorkloadCount + testWorkloadCount + 2 + 2 + 10 + 10;
        }

        #endregion workload
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
