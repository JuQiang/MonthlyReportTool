﻿using System;
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

        public void Build(string project)
        {
            BuildTitle();
            BuildSubTitle();
            BuildDescription();

            BuildSummaryTable(4);
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
        private void BuildSummaryTable(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(sheet, startRow, "本迭代产品特性完成情况统计", "说明：统计依据为：完成本月目标状态的本月目标日期落在本迭代期间内的产品特性", "B", "O",
                new List<string>() { "分类", "个数", "占比", "说明"},
                new List<string>() {"B,D","E,E","F,F","G,P" },
                6);

            string[] cols1 = new string[] { "已完成数", "拖期数","中止数","移除数", "按计划完成数", "本迭代计划总数" };
            string[] cols2 = new string[] { "=IF(E7<>0,E7/E12,\"\")", "=IF(E8<>0,E8/E12,\"\")","=IF(E9<>0,E9/E12,\"\")",
                "=IF(E10<>0,E10/E12,\"\")","=IF(E11<>0,E11/E12,\"\")","'--"

            };
            string[] cols3 = new string[] { "已完成数：已完成本月目标的产品特性个数\r\n占比：已完成数/本迭代计划总数",
                "拖期数：未完成本迭代目标的产品特性个数\r\n占比：拖期数/本迭代计划总数",
                "中止数：本迭代中止的产品特性个数\r\n占比：移除数/本迭代计划总数",
                "移除数：本迭代移除的产品特性个数\r\n占比：移除数/本迭代计划总数",
                "按计划完成数：按本月目标日期完成的产品特性个数\r\n占比：按计划完成数/本迭代计划总数",
                "本迭代时间范围内所有迭代产品特性总数本迭代计划总数=已完成数+拖期数" };

            for (int row = 7; row <= 12; row++)
            {
                sheet.Cells[row, "B"] = cols1[row - 7];
                sheet.Cells[row, "F"] = cols2[row - 7];
                sheet.Cells[row, "G"] = cols3[row - 7];
            }
            
            sheet.Cells[12, "E"] = "=SUM(E7: E11)";

            ExcelInterop.Range range = sheet.Range[sheet.Cells[7, "B"], sheet.Cells[12, "B"]];
            Utility.AddNativieResource(range);
            range.RowHeight = 40;
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
