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
using System.Drawing;

namespace MonthlyReportTool.API.Office.Excel
{
    public class Utility
    {
        private static List<object> nativeResources = new List<object>();
        public static void AddNativieResource(object obj)
        {
            nativeResources.Add(obj);
        }
        public static void BuildIterationReport(ProjectEntity project, string path)
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
                Tuple.Create<string, Type>("项目整体说明",typeof(OverviewSheet)),
                Tuple.Create<string, Type>("系统需求统计分析",typeof(FeatureSheet)),
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

                Console.WriteLine("\t" + sheet.Name);

                Type t = allSheets[i].Item2;
                var ci = t.GetConstructor(new Type[] { typeof(ExcelInterop.Worksheet) });
                object obj = ci.Invoke(new object[] { sheet });
                t.InvokeMember("Build", BindingFlags.InvokeMethod, null, obj, new object[] { project });
            }

            Console.WriteLine("保存文件.");
            sheets.Select();//选择所有的sheet

            var window = excel.ActiveWindow;
            nativeResources.Add(window);
            window.DisplayGridlines = false;//都不显示表格线

            var ite = TFS.Utility.GetBestIteration(project.Name);
            string fname = String.Format("{0}\\{1}总结({2}_{3}).xlsx",
                path,
                ite.Path.Replace("\\", " "),
                (DateTime.Parse(ite.StartDate)).ToString("yyyyMMdd"),
                (DateTime.Parse(ite.EndDate)).ToString("yyyyMMdd")
            );
            workbook.SaveAs(fname);
            workbook.Close();

            Console.WriteLine("释放资源.");
            foreach (object com in nativeResources)
            {
                TFS.Utility.ReleaseComObject(com);
            }
            Console.WriteLine("附件已经保存在：" + fname);

            excel.Quit();
        }

        public static void SetFormatBigger(ExcelInterop.Range range, double limit)
        {
            Utility.AddNativieResource(range);
            ExcelInterop.FormatConditions formcond = range.FormatConditions;
            Utility.AddNativieResource(formcond);
            ExcelInterop.FormatCondition newcond = formcond.Add(ExcelInterop.XlFormatConditionType.xlCellValue, ExcelInterop.XlFormatConditionOperator.xlGreaterEqual, limit);
            Utility.AddNativieResource(newcond);
            ExcelInterop.Font condfont = newcond.Font;
            Utility.AddNativieResource(condfont);
            condfont.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); //Red letters
        }

        public static void SetFormatSmaller(ExcelInterop.Range range, double limit)
        {
            Utility.AddNativieResource(range);
            ExcelInterop.FormatConditions formcond = range.FormatConditions;
            Utility.AddNativieResource(formcond);
            ExcelInterop.FormatCondition newcond = formcond.Add(ExcelInterop.XlFormatConditionType.xlCellValue, ExcelInterop.XlFormatConditionOperator.xlLess, limit);
            Utility.AddNativieResource(newcond);
            ExcelInterop.Font condfont = newcond.Font;
            Utility.AddNativieResource(condfont);
            condfont.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); //Red letters
        }
        public static string GetPersonName(string fullname)
        {
            if (fullname.Trim().Length < 1) return "";
            return fullname.Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();

        }
        public static void SetCellPercentFormat(ExcelInterop.Range range)
        {
            Utility.AddNativieResource(range);
            range.NumberFormat = "0%";
        }
        public static void SetCellPercentFormat(ExcelInterop.Worksheet sheet, int startRow, int startCol, int endRow, int endCol)
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

        public static void SetCellColor(ExcelInterop.Range range, System.Drawing.Color color)
        {
            var rangeFont = range.Font;
            Utility.AddNativieResource(rangeFont);
            rangeFont.Color = ColorTranslator.ToOle(color);//是OLE的颜色，不是GDI+的颜色。
        }

        public static void SetCellBorder(ExcelInterop.Range range)
        {
            Utility.AddNativieResource(range);
            var borders = range.Borders;
            Utility.AddNativieResource(borders);
            borders.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
        }


        public static void SetCellColor(ExcelInterop.Range range, System.Drawing.Color color, string text, bool bold = true)
        {
            string fulltext = Convert.ToString(range.Text);
            int pos = fulltext.IndexOf(text);

            if (pos < 0) return;

            var tmpchar = range.Characters[pos + 1, text.Length];
            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = bold;
            tmpfont.Color = ColorTranslator.ToOle(color);//是OLE的颜色，不是GDI+的颜色。
        }

        public static void SetCellFontRedColor(ExcelInterop.Range range)
        {
            SetCellColor(range, System.Drawing.Color.Red);
        }
        public static void SetCellDarkGrayColor(ExcelInterop.Range range)
        {
            var interior = range.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();
        }
        public static void SetCellGreenColor(ExcelInterop.Range range)
        {
            var interior = range.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.FromArgb(205, 255, 202);// System.Drawing.Color.Green.ToArgb();
            //#CAFFCE
        }

        public static void SetCellAlignAndWrap(ExcelInterop.Range range, ExcelInterop.XlHAlign hAlign = ExcelInterop.XlHAlign.xlHAlignLeft)
        {
            Utility.AddNativieResource(range);
            range.HorizontalAlignment = hAlign;
            range.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            range.WrapText = true;
        }
        public static int BuildFormalTable(ExcelInterop.Worksheet sheet, int row, string title, string description,
            string startCol, string endCol, List<string> colnames, List<string> mergedInfo, int rowCount, bool nodata = false)
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

            if (!nodata)
            {
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

                SetTableHeaderFormat(sheet.get_Range(String.Format("{0}{1}:{2}{3}", startCol, row - 1, endCol, row - 1)));
            }
            else
            {
                ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[row, startCol], sheet.Cells[row + rowCount, endCol]];
                Utility.AddNativieResource(tableRange);
                tableRange.Merge();
                tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignLeft;
                tableRange.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignTop;

                var interior = tableRange.Interior;
                Utility.AddNativieResource(interior);
                interior.Color = System.Drawing.Color.White.ToArgb();

                var border = tableRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            }
            return row + rowCount + 2;
        }

        public static void SetTableHeaderFormat(ExcelInterop.Range range, bool bold = true)
        {
            SetCellAlignAndWrap(range);
            Utility.AddNativieResource(range);

            SetCellBorder(range);

            var interior = range.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var font = range.Font;
            Utility.AddNativieResource(font);
            font.Bold = bold;

            Utility.SetCellAlignAndWrap(range, ExcelInterop.XlHAlign.xlHAlignCenter);
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
    }
}
