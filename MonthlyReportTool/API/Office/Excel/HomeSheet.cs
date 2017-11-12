using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.Excel
{
    public class HomeSheet : ExcelSheetBase,IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public HomeSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            ExcelInterop.Range allrange = sheet.Range[sheet.Cells[1, "C"], sheet.Cells[40, "J"]];
            Utility.AddNativieResource(allrange);
            allrange.ColumnWidth = 12;
            allrange.RowHeight = 15;

            #region 1st paragraph
            sheet.Cells[5, "D"] = "***迭代总结";
            ExcelInterop.Range range = sheet.Range[sheet.Cells[5, "D"], sheet.Cells[6, "H"]];
            Utility.AddNativieResource(range);
            range.Merge();
            var font = range.Font;
            font.Size = 20;
            font.Name = "微软雅黑";
            font.Bold = true;
            Utility.AddNativieResource(font);
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
            Utility.AddNativieResource(range2);
            range2.Merge();
            range2.UseStandardHeight = true;
            range2.WrapText = true;
            range2.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            var font2 = range2.Font;
            Utility.AddNativieResource(font2);
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

                Utility.AddNativieResource(tmpfont);
                Utility.AddNativieResource(tmpcharc);
            }

            var border2 = range2.Borders;
            Utility.AddNativieResource(border2);
            border2.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 2nd paragraph

            #region 3rd paragraph
            text = "表格内容使用说明：1、表格中需要公式计算的地方已添加公式，填入统计数据后，派生数据会自动生成，底色为绿色的不要随意改动；";
            sheet.Cells[19, "C"] = text;

            ExcelInterop.Range range3 = sheet.Range[sheet.Cells[19, "C"], sheet.Cells[23, "J"]];
            Utility.AddNativieResource(range3);
            range3.Merge();
            range3.UseStandardHeight = true;
            range3.WrapText = true;
            range3.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            var font3 = range3.Font;
            Utility.AddNativieResource(font3);
            font3.Size = 11;
            font3.Name = "微软雅黑";

            var tmpcharc3 = range3.Characters[1, 9];
            var tmpfont3 = tmpcharc3.Font;
            tmpfont3.Bold = true;

            Utility.AddNativieResource(tmpfont3);
            Utility.AddNativieResource(tmpcharc3);


            var border3 = range3.Borders;
            Utility.AddNativieResource(border3);
            border3.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 3rd paragraph

            #region 4th paragraph
            text = "项目模板定义：1、此模板为组织级模板，各团队项目根据人员投入及团队项目情况定义自己的模板，团队项目模板生成后，每次直接填写数据即可，不必每次调整模板格式；";
            sheet.Cells[25, "C"] = text;

            ExcelInterop.Range range4 = sheet.Range[sheet.Cells[25, "C"], sheet.Cells[28, "J"]];
            Utility.AddNativieResource(range4);
            range4.Merge();
            range4.UseStandardHeight = true;
            range4.WrapText = true;
            range4.VerticalAlignment = ExcelInterop.XlVAlign.xlVAlignCenter;
            var font4 = range4.Font;
            Utility.AddNativieResource(font4);
            font4.Size = 11;
            font4.Name = "微软雅黑";

            var tmpcharc4 = range4.Characters[1, 9];
            var tmpfont4 = tmpcharc4.Font;
            tmpfont4.Bold = true;

            Utility.AddNativieResource(tmpfont4);
            Utility.AddNativieResource(tmpcharc4);


            var border4 = range4.Borders;
            Utility.AddNativieResource(border4);
            border4.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            #endregion 3rd paragraph
        }
    }
}
