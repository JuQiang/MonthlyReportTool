using System;
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

        public static void SetSheetFont(ExcelInterop.Worksheet sheet)
        {
            var bigrange = sheet.Range[sheet.Cells[1, "A"], sheet.Cells[1000, "Z"]];
            nativeResources.Add(bigrange);
            var bigrangeFont = bigrange.Font;
            nativeResources.Add(bigrangeFont);

            bigrangeFont.Name = "微软雅黑";
            bigrangeFont.Size = 11;
        }
       
        private static void BuildSuggestionSheet(ExcelInterop.Worksheet sheet)
        {

        }

        private static void BuildPeoplePerformanceSheet(ExcelInterop.Worksheet sheet)
        {

        }

    }
}
