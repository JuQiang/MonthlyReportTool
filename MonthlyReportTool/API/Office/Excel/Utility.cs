using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using MonthlyReportTool.API.TFS.TeamProject;
using MonthlyReportTool.API.TFS.Agile;
using System.Diagnostics;

namespace MonthlyReportTool.API.Office.Excel
{
    public class Utility
    {
        private static List<object> nativeResources = new List<object>();
        public static void AddNativieResource(object obj)
        {
            nativeResources.Add(obj);
        }
        public static void BuildIterationReports(ProjectEntity project)
        {
            ExcelInterop.Application excel = new ExcelInterop.Application();

            excel.DisplayAlerts = false;

            ExcelInterop.Workbook workbook = excel.Workbooks.Add();
            nativeResources.Add(workbook);
            ExcelInterop.Sheets sheets = workbook.Worksheets;
            nativeResources.Add(sheets);

            List<Tuple<string, Type>> allSheets = new List<Tuple<string, Type>>()
            {
                Tuple.Create<string, Type>("首页及说明",typeof(HomeSheet)),
                //Tuple.Create<string, Type>("目录",typeof(ContentSheet)),
                Tuple.Create<string, Type>("项目整体说明",typeof(OverviewSheet)),
                Tuple.Create<string, Type>("产品特性统计",typeof(FeatureSheet)),
                Tuple.Create<string, Type>("Backlog统计",typeof(BacklogSheet)),
                Tuple.Create<string, Type>("工作量统计",typeof(WorkloadSheet)),
                Tuple.Create<string, Type>("提交单分析",typeof(CommitmentSheet)),
                Tuple.Create<string, Type>("代码审查分析",typeof(CodeReviewSheet)),
                Tuple.Create<string, Type>("Bug统计分析",typeof(BugSheet)),
                Tuple.Create<string, Type>("改进建议",typeof(SuggestionSheet)),
                Tuple.Create<string, Type>("人员考评结果",typeof(PerformanceSheet)),
            };

            ExcelInterop.Worksheet lastSheet = null;
            for (int i = 0; i < allSheets.Count; i++)
            {
                ExcelInterop.Worksheet sheet;

                if (lastSheet == null) sheet = (ExcelInterop.Worksheet)sheets.Add(); else sheet = (ExcelInterop.Worksheet)sheets.Add(After: lastSheet);

                lastSheet = sheet;
                nativeResources.Add(sheet);
                sheet.Name = allSheets[i].Item1;

                WriteLog("\t" + sheet.Name);

                Type t = allSheets[i].Item2;
                var ci = t.GetConstructor(new Type[] { typeof(ExcelInterop.Worksheet) });
                object obj = ci.Invoke(new object[] { sheet });
                t.InvokeMember("Build", BindingFlags.InvokeMethod, null, obj, new object[] { project });
            }

            WriteLog("保存文件.");
            sheets.Select();//选择所有的sheet

            var window = excel.ActiveWindow;
            nativeResources.Add(window);
            window.DisplayGridlines = false;//都不显示表格线

            var ite = TFS.Utility.GetBestIteration(project.Name);
            workbook.SaveAs(String.Format("c:\\irt\\{0}总结({1}_{2}).xlsx",
                ite.Path.Replace("\\", " "),
                (DateTime.Parse(ite.StartDate)).ToString("yyyyMMdd"),
                (DateTime.Parse(ite.EndDate)).ToString("yyyyMMdd")
                ));
            workbook.Close();

            WriteLog("释放资源.");
            foreach (object com in nativeResources)
            {
                TFS.Utility.ReleaseComObject(com);
            }

            excel.Quit();
        }

        public static string GetPersonName(string fullname)
        {
            if (fullname.Trim().Length < 1) return "";
            return fullname.Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();

        }
        public static void SetupSheetPercentFormat(ExcelInterop.Worksheet sheet, int startRow, string startCol, int endRow, string endCol)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];
            Utility.AddNativieResource(range);
            range.NumberFormat = "0%";
        }
        public static void SetupSheetPercentFormat(ExcelInterop.Worksheet sheet, int startRow, int startCol, int endRow, int endCol)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];
            Utility.AddNativieResource(range);
            range.NumberFormat = "0%";
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
            string startCol, string endCol, List<string> colnames, List<string> mergedInfo, int rowCount)
        {
            //Utility.WriteLog("Build Formal Table - Begin.");
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
            var lines = description.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            tableDescriptionRange.RowHeight = 20 * (lines.Length + 0);

            row++;

            //标题栏
            for (int i = 0; i < colnames.Count; i++)
            {
                string[] cols = mergedInfo[i].Split(new char[] { ',' });
                sheet.Cells[row, cols[0]] = colnames[i];//table header
            }
            StringBuilder sb = new StringBuilder();
            for (int j = 0; j < colnames.Count; j++)
            {
                string[] cols = mergedInfo[j].Split(new char[] { ',' });
                if (cols[0].ToLower() == cols[1].ToLower()) continue;

                sb.AppendFormat("{0}{1}:{2}{3},", cols[0], row, cols[1], row);
            }

            if (sb.Length > 0)
            {
                sb.Remove(sb.Length - 1, 1);

                ExcelInterop.Range colRange = sheet.get_Range(sb.ToString());
                Utility.AddNativieResource(colRange);
                colRange.Merge();

            }
            
            ExcelInterop.Range firstRow = sheet.get_Range(String.Format("{0}{1}:{2}{3}", startCol, row, endCol, row));
            Utility.AddNativieResource(firstRow);
            firstRow.RowHeight = 20;
            var border = firstRow.Borders;
            Utility.AddNativieResource(border);
            border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            firstRow.Copy();
            row++;
            for (int i = 0; i < rowCount; i++)
            {
                var destRange = sheet.get_Range(String.Format("{0}{1}:{2}{3}", startCol, row + i, endCol, row + i));
                Utility.AddNativieResource(destRange);
                destRange.PasteSpecial(ExcelInterop.XlPasteType.xlPasteFormats);
                //firstRow.Copy(startCol + ":" + row + i);
            }

            SetTableHeaderFormat(sheet, row-1, startCol, row-1, endCol);
            
            return row + rowCount + 2;
        }

        public static void SetTableHeaderFormat(ExcelInterop.Worksheet sheet, int startRow, string startCol, int endRow, string endCol)
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

        public static void BuildFormalSheetTitle(ExcelInterop.Worksheet sheet, int startRow, string startCol, int endRow, string endCol, string title, int columnWidth = 16)
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
            Console.WriteLine(msg);

            using (StreamWriter sw = new StreamWriter(@"c:\\irt\\log.txt", true))
            {
                sw.WriteLine(line);
            }
        }

    }
}
