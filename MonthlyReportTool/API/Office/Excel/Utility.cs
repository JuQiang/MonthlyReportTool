using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;

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
            WriteLog("================开始================");
            ExcelInterop.Application excel = new ExcelInterop.Application();

            excel.DisplayAlerts = false;

            ExcelInterop.Workbook workbook = excel.Workbooks.Add();
            nativeResources.Add(workbook);
            ExcelInterop.Sheets sheets = workbook.Worksheets;
            nativeResources.Add(sheets);

            List<Tuple<string, Type>> allSheets = new List<Tuple<string, Type>>()
            {
                Tuple.Create<string, Type>("首页及说明",typeof(HomeSheet)),
                Tuple.Create<string, Type>("目录",typeof(ContentSheet)),
                Tuple.Create<string, Type>("项目整体说明",typeof(OverviewSheet)),
                Tuple.Create<string, Type>("产品特性统计",typeof(FeatureSheet)),
                Tuple.Create<string, Type>("Backlog统计",typeof(BacklogSheet)),
                Tuple.Create<string, Type>("工作量统计",typeof(WorkloadSheet)),
                Tuple.Create<string, Type>("提交单分析",typeof(CommitmentSheet)),
                Tuple.Create<string, Type>("代码审查分析",typeof(CodeReviewSheet)),
                Tuple.Create<string, Type>("Bug统计分析",typeof(BugAnalysisSheet)),
                Tuple.Create<string, Type>("改进建议",typeof(SuggestionSheet)),
                Tuple.Create<string, Type>("人员考评结果",typeof(PerformanceSheet)),
            };

            ExcelInterop.Worksheet lastSheet = null;
            for (int i = 0; i < allSheets.Count; i++)
            {
                ExcelInterop.Worksheet sheet;

                if (lastSheet == null)sheet = (ExcelInterop.Worksheet)sheets.Add();else sheet = (ExcelInterop.Worksheet)sheets.Add(After: lastSheet);

                lastSheet = sheet;
                nativeResources.Add(sheet);
                sheet.Name = allSheets[i].Item1;

                WriteLog("处理：" + sheet.Name);

                Type t = allSheets[i].Item2;
                var ci = t.GetConstructor(new Type[] { typeof(ExcelInterop.Worksheet) });
                object obj = ci.Invoke(new object[] { sheet });
                t.InvokeMember("Build", BindingFlags.InvokeMethod, null, obj, new object[] { });
            }

            WriteLog("保存文件.");
            sheets.Select();//选择所有的sheet

            var window = excel.ActiveWindow;
            nativeResources.Add(window);
            window.DisplayGridlines = false;//都不显示表格线

            workbook.SaveAs("c:\\irt\\1.xlsx");
            workbook.Close();

            WriteLog("释放资源.");
            foreach (object com in nativeResources)
            {
                TFS.Utils.ReleaseComObject(com);
            }

            excel.Quit();

            WriteLog("================结束================");
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

        public static void WriteLog(string msg)
        {
            string line = String.Format("{0} --- {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), msg);
            Console.WriteLine(line);

            using (StreamWriter sw = new StreamWriter(@"c:\\irt\\log.txt",true))
            {
                sw.WriteLine(line);
            }
        }
    }
}
