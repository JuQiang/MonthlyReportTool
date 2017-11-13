﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;

namespace MonthlyReportTool.API.Office.Excel
{
    public class Utility
    {
        private static List<object> nativeResources = new List<object>();
        public static void AddNativieResource(object obj)
        {
            nativeResources.Add(obj);
        }
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
            (new HomeSheet(sheetHome)).Build();

            var sheetContent = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetHome);
            nativeResources.Add(sheetContent);
            sheetContent.Name = "目录";
            (new HomeSheet(sheetContent)).Build();

            var sheetOverview = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetContent);
            nativeResources.Add(sheetOverview);
            sheetOverview.Name = "项目整体说明";
            (new OverviewSheet(sheetOverview)).Build();

            var sheetFeature = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetOverview);
            nativeResources.Add(sheetFeature);
            sheetFeature.Name = "产品特性统计";
            (new FeatureSheet(sheetFeature)).Build();

            var sheetBacklog = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetFeature);
            nativeResources.Add(sheetBacklog);
            sheetBacklog.Name = "Backlog统计";
            (new BacklogSheet(sheetBacklog)).Build();

            var sheetWorkload = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetBacklog);
            nativeResources.Add(sheetWorkload);
            sheetWorkload.Name = "工作量统计";
            (new WorkloadSheet(sheetWorkload)).Build();

            var sheetCommitment = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetWorkload);
            nativeResources.Add(sheetCommitment);
            sheetCommitment.Name = "提交单分析";
            (new CommitmentSheet(sheetCommitment)).Build();
            
            var sheetCodeReview = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetCommitment);
            nativeResources.Add(sheetCodeReview);
            sheetCodeReview.Name = "代码审查分析";
            (new CodeReviewSheet(sheetCodeReview)).Build();

            var sheetBugAnalysis = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetCodeReview);
            nativeResources.Add(sheetBugAnalysis);
            sheetBugAnalysis.Name = "Bug统计分析";
            (new BugAnalysisSheet(sheetBugAnalysis)).Build();

            var sheetSuggestion = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetBugAnalysis);
            nativeResources.Add(sheetSuggestion);
            sheetSuggestion.Name = "改进建议";
            (new SuggestionSheet(sheetSuggestion)).Build();

            var sheetPeoplePerformance = (ExcelInterop.Worksheet)sheets.Add(Missing.Value, sheetSuggestion);
            nativeResources.Add(sheetPeoplePerformance);
            sheetPeoplePerformance.Name = "人员考评结果";
            (new PerformanceSheet(sheetPeoplePerformance)).Build();

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

        public static void SetSheetFont(ExcelInterop.Worksheet sheet)
        {
            var bigrange = sheet.Range[sheet.Cells[1, "A"], sheet.Cells[1000, "Z"]];
            nativeResources.Add(bigrange);
            var bigrangeFont = bigrange.Font;
            nativeResources.Add(bigrangeFont);

            bigrangeFont.Name = "微软雅黑";
            bigrangeFont.Size = 11;
        }

        public static int BuildFormalTable(ExcelInterop.Worksheet sheet, int row, string title, string description,
            string startCol, string endCol, List<string> colnames, List<string> mergedInfo,int rowCount)
        {
            ExcelInterop.Range tableTitleRange = sheet.Range[sheet.Cells[row, startCol], sheet.Cells[row, endCol]];
            Utility.AddNativieResource(tableTitleRange);
            tableTitleRange.Merge();
            tableTitleRange.RowHeight = 20;
            sheet.Cells[row, startCol] = title;
            var tableTitleFont = tableTitleRange.Font;
            Utility.AddNativieResource(tableTitleFont);
            tableTitleFont.Bold = true;
            tableTitleFont.Size = 12;

            row++;

            ExcelInterop.Range tableDescriptionRange = sheet.Range[sheet.Cells[row, startCol], sheet.Cells[row, endCol]];
            Utility.AddNativieResource(tableDescriptionRange);
            tableDescriptionRange.Merge();
            
            sheet.Cells[row, startCol] = description;
            var lines = description.Split(new char[] { '\r', '\n' },StringSplitOptions.RemoveEmptyEntries);
            tableDescriptionRange.RowHeight = 20*(lines.Length+0);

            row++;
            for (int i = 0; i < colnames.Count; i++)
            {
                string[] cols = mergedInfo[i].Split(new char[] { ',' });
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row,cols[0]], sheet.Cells[row, cols[1]]];
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 20;
                colRange.Merge();
                sheet.Cells[row, cols[0]] = colnames[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            BuildFormalTableHeader(sheet, row, startCol, row, endCol);

            row++;
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colnames.Count; j++)
                {
                    string[] cols = mergedInfo[j].Split(new char[] { ',' });
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[row + i, cols[0]], sheet.Cells[row + i, cols[1]]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();

                    sheet.Cells[row + i, cols[0]] = String.Format("数据行:{0}，列{1}", row + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

            return row + rowCount + 1;
        }

        public static void BuildFormalTableHeader(ExcelInterop.Worksheet sheet,int startRow, string startCol, int endRow, string endCol)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];
            Utility.AddNativieResource(range);
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            range.WrapText = true;

            var borders = range.Borders;
            Utility.AddNativieResource(borders);
            borders.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            var interior = range.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();
        }

        public static void BuildFormalSheetTitle(ExcelInterop.Worksheet sheet, int startRow, string startCol, int endRow, string endCol,string title, int columnWidth=16)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];

            Utility.AddNativieResource(range);
            range.ColumnWidth = columnWidth;
            range.RowHeight = 40;
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;
            range.Merge();
            sheet.Cells[startRow, startCol] = title;
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 20;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
        }
    }
}
