using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;

namespace MonthlyReportTool.API.Office.Excel
{
    public class FeatureSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        public FeatureSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build()
        {
            BuildTitle();
            BuildSubTitle();
            BuildDescription();

            BuildSummaryTable();
            List<FeatureEntity> list = new List<FeatureEntity>() { new FeatureEntity(), new FeatureEntity(), new FeatureEntity(), new FeatureEntity(), new FeatureEntity() };

            int startRow = BuildDelayTable(14, list);
            startRow = BuildAnadonTable(startRow,list);
            startRow = BuildTable(startRow, list);
        }

        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "O", "产品特性统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "本迭代产品特性完成情况统计";
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
            sheet.Cells[5, "B"] = "说明：统计依据为：完成本月目标状态的本月目标日期落在本迭代期间内的产品特性";

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
                Utility.AddNativieResource(colRange);
                colRange.RowHeight = 40;
                colRange.Merge();
                sheet.Cells[row, "B"] = cols[0, row - 6];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                ExcelInterop.Range col2Range = sheet.Cells[row, "E"] as ExcelInterop.Range;
                Utility.AddNativieResource(col2Range);
                col2Range.RowHeight = 40;
                col2Range.Merge();
                sheet.Cells[row, "E"] = cols[1, row - 6];

                var border2 = col2Range.Borders;
                Utility.AddNativieResource(border2);
                border2.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                ExcelInterop.Range col3Range = sheet.Cells[row, "F"] as ExcelInterop.Range;
                Utility.AddNativieResource(col3Range);
                col3Range.RowHeight = 40;
                col3Range.Merge();
                sheet.Cells[row, "F"] = cols[2, row - 6];

                var border3 = col3Range.Borders;
                Utility.AddNativieResource(border3);
                border3.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

                ExcelInterop.Range col4Range = sheet.Range[sheet.Cells[row, "G"], sheet.Cells[row, "O"]];
                Utility.AddNativieResource(col4Range);
                col4Range.RowHeight = 40;
                col4Range.Merge();
                sheet.Cells[row, "G"] = cols[3, row - 6];

                var border4 = col4Range.Borders;
                Utility.AddNativieResource(border4);
                border4.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range col5Range = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "O"]];
            Utility.AddNativieResource(col5Range);
            col5Range.RowHeight = 20;
            col5Range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = col5Range.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            var col5Font = col5Range.Font;
            Utility.AddNativieResource(col5Font);
            col5Font.Name = "微软雅黑";
            col5Font.Size = 11;
            col5Font.Bold = true;

            sheet.Cells[7, "F"] = "=IF(E7<>0,E7/E10,\"\")";
            sheet.Cells[8, "F"] = "=IF(E8<>0,E8/E10,\"\")";
            sheet.Cells[9, "F"] = "=IF(E9<>0,E9/E10,\"\")";
            sheet.Cells[10, "E"] = "=SUM(E7: E8)";
            sheet.Cells[10, "F"] = "'--";
        }

        private int BuildDelayTable(int startRow,List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代拖期产品特性分析", "说明：按关键应用、模块排序；非研发类的为无", "B", "P",
                new List<string>() { "ID", "关键应用", "模块", "产品特性名称", "负责人", "拖期原因分析" },
                new List<string>() { "B,B", "C,D", "E,F", "G,J", "K,L", "M,P" },
                features.Count);
            return nextRow;
        }
        private int BuildAnadonTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代移除/中止产品特性分析", "说明：按关键应用、模块排序；非研发类的为无", "B", "O",
                new List<string>() { "ID", "关键应用", "模块", "产品特性名称", "负责人", "移除/中止原因说明" },
                new List<string>() { "B,B", "C,D", "E,F", "G,J", "K,L", "M,P" },
                features.Count);

            return nextRow;
        }
        private int BuildTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代产品特性列表", "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      如果本迭代实际的产品特性数多于模板预制的行数，请自行插入行，然后用格式刷刷新增的行的格式\r\n      按关键应用、模块排序；非研发类的为无", "B", "O",
                new List<string>() { "ID", "关键应用", "模块", "产品特性名称", "目标状态", "目标日期", "负责人", "当前状态" },
                new List<string>() { "B,B", "C,D", "E,F", "G,I", "J,K", "L,M","N,N","O,O" },
                features.Count);

            return nextRow;
            
        }
    }
}
