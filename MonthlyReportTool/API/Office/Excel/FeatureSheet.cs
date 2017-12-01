using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.TeamProject;

namespace MonthlyReportTool.API.Office.Excel
{
    public class FeatureSheet : ExcelSheetBase, IExcelSheet
    {
        private List<List<FeatureEntity>> featureList;
        private ExcelInterop.Worksheet sheet;
        public FeatureSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            this.featureList = Feature.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));
            BuildTitle();
            BuildSubTitle();
            BuildDescription();

            BuildSummaryTable(4);


            int startRow = BuildTable(13, this.featureList[0]);

            startRow = BuildAnadonTable(startRow, this.featureList[2]);
            startRow = BuildDelayTable(startRow, this.featureList[3]);

            var colKL = sheet.get_Range("K1:L1");
            Utility.AddNativieResource(colKL);
            colKL.ColumnWidth = 6.27d;

            sheet.Cells[1, "A"] = "";
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
                new List<string>() { "分类", "个数", "占比", "说明" },
                new List<string>() { "B,D", "E,E", "F,F", "G,P" },
                5);

            string[] cols1 = new string[] { "已完成数", "拖期数", "中止/移除数", "按计划完成数", "本迭代计划总数" };
            string[] cols2 = new string[] { "=IF(E11<>0,E7/E11,\"\")", "=IF(E11<>0,E8/E11,\"\")", "'--", "=IF(E11<>0,E9/E11,\"\")", "'--" };
            string[] cols3 = new string[] { "已完成数：已完成本迭代目标的产品特性个数\r\n占比：已完成数/本迭代计划总数",
                "拖期数：未完成本迭代目标的产品特性个数\r\n占比：拖期数/本迭代计划总数",
                "中止/移除数：本迭代期间内中止/移除的产品特性个数\r\n占比：移除数/本迭代计划总数",
                "按计划完成数：按本迭代目标日期完成的产品特性个数\r\n占比：按计划完成数/本迭代计划总数",
                "本迭代内的所有产品特性总数（本迭代目标日期在本迭代期间内的所有）" };

            for (int row = 7; row <= 11; row++)
            {
                sheet.Cells[row, "B"] = cols1[row - 7];
                sheet.Cells[row, "F"] = cols2[row - 7];
                sheet.Cells[row, "G"] = cols3[row - 7];
            }

            sheet.Cells[7, "E"] = this.featureList[1].Count;
            sheet.Cells[8, "E"] = this.featureList[3].Count;
            sheet.Cells[9, "E"] = this.featureList[2].Count;
            sheet.Cells[10, "E"] = this.featureList[4].Count;
            sheet.Cells[11, "E"] = this.featureList[0].Count;
            //sheet.Cells[12, "E"] = "=SUM(E7: E11)";

            Utility.SetCellPercentFormat(sheet.get_Range("F7:F11"));

            ExcelInterop.Range range = sheet.Range[sheet.Cells[7, "B"], sheet.Cells[12, "B"]];
            Utility.AddNativieResource(range);
            range.RowHeight = 40;

            Utility.SetFormatBigger(sheet.Cells[8, "E"], 0.0001d);
            Utility.SetFormatBigger(sheet.Cells[9, "E"], 0.0001d);
        }

        private int BuildDelayTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代拖期产品特性分析", "说明：按关键应用、模块排序；非研发类的为无", "B", "T",
                new List<string>() { "ID", "关键应用", "模块", "产品特性名称", "本迭代目标状态", "当前状态", "迭代目标日期","本月目标日期", "负责人", "移除/中止原因说明" },
                new List<string>() { "B,B", "C,D", "E,F", "G,J", "K,L", "M,M", "N,N", "O,O","P,P", "Q,T" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块排序");
            var orderedFeatures = features.OrderBy(feature => feature.KeyApplication).ThenBy(feature => feature.ModulesName).ToList();
            startRow += 3;
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = orderedFeatures[i].Id;
                sheet.Cells[startRow + i, "C"] = orderedFeatures[i].KeyApplication;
                sheet.Cells[startRow + i, "E"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "G"] = orderedFeatures[i].Title;
                sheet.Cells[startRow + i, "K"] = orderedFeatures[i].MonthState;
                sheet.Cells[startRow + i, "M"] = orderedFeatures[i].State;
                sheet.Cells[startRow + i, "N"] = DateTime.Parse(orderedFeatures[i].IterationTargetDate).AddHours(8).ToString("yyyy-MM-dd");
                sheet.Cells[startRow + i, "O"] = DateTime.Parse(orderedFeatures[i].TargetDate).AddHours(8).ToString("yyyy-MM-dd");
                sheet.Cells[startRow + i, "P"] = Utility.GetPersonName(orderedFeatures[i].AssignedTo);
                sheet.Cells[startRow + i, "Q"] = "";
            }

            Utility.SetCellRedColor(sheet.Cells[startRow - 1, "Q"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);

            return nextRow;
        }
        private int BuildAnadonTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代移除/中止产品特性分析", "说明：按关键应用、模块排序；非研发类的为无", "B", "T",
                new List<string>() { "ID", "关键应用", "模块", "产品特性名称", "本迭代目标状态", "当前状态","迭代目标日期", "本月目标日期", "负责人", "移除/中止原因说明" },
                new List<string>() { "B,B", "C,D", "E,F", "G,J", "K,L", "M,M","N,N","O,O","P,P","Q,T" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[startRow+1, "B"], System.Drawing.Color.Red, "按关键应用、模块排序");
            var orderedFeatures = features.OrderBy(feature => feature.KeyApplication).ThenBy(feature => feature.ModulesName).ToList();
            startRow += 3;
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = orderedFeatures[i].Id;
                sheet.Cells[startRow + i, "C"] = orderedFeatures[i].KeyApplication;
                sheet.Cells[startRow + i, "E"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "G"] = orderedFeatures[i].Title;
                sheet.Cells[startRow + i, "K"] = orderedFeatures[i].MonthState;
                sheet.Cells[startRow + i, "M"] = orderedFeatures[i].State;
                sheet.Cells[startRow + i, "N"] = DateTime.Parse(orderedFeatures[i].IterationTargetDate).AddHours(8).ToString("yyyy-MM-dd");
                sheet.Cells[startRow + i, "O"] = DateTime.Parse(orderedFeatures[i].TargetDate).AddHours(8).ToString("yyyy-MM-dd");
                sheet.Cells[startRow + i, "P"] = Utility.GetPersonName(orderedFeatures[i].AssignedTo);
                sheet.Cells[startRow + i, "Q"] = "";
            }

            Utility.SetCellRedColor(sheet.Cells[startRow - 1, "Q"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);
            return nextRow;
        }
        private int BuildTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代产品特性列表", "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      如果本迭代实际的产品特性数多于模板预制的行数，请自行插入行，然后用格式刷刷新增的行的格式\r\n      按关键应用、模块排序；非研发类的为无", "B", "P",
                new List<string>() { "ID", "关键应用", "模块", "产品特性名称", "本迭代目标状态", "当前状态", "迭代目标日期","本月目标日期", "负责人"},
                new List<string>() { "B,B", "C,D", "E,F", "G,J", "K,L", "M,M", "N,N", "O,O","P,P" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[14, "B"], System.Drawing.Color.Red, "按关键应用、模块排序");

            var orderedFeatures = features.OrderBy(feature => feature.KeyApplication).ThenBy(feature => feature.ModulesName).ToList();
            startRow += 3;
            object[,] arr = new object[orderedFeatures.Count, 15];
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                arr[i, 0] = orderedFeatures[i].Id;
                arr[i, 1] = orderedFeatures[i].KeyApplication;
                arr[i, 3] = orderedFeatures[i].ModulesName;
                arr[i, 5] = orderedFeatures[i].Title;
                arr[i, 9] = orderedFeatures[i].MonthState;
                arr[i, 11] = orderedFeatures[i].State;
                arr[i, 12] = DateTime.Parse(orderedFeatures[i].IterationTargetDate).AddHours(8).ToString("yyyy-MM-dd");
                arr[i, 13] = DateTime.Parse(orderedFeatures[i].TargetDate).AddHours(8).ToString("yyyy-MM-dd");
                arr[i, 14] = Utility.GetPersonName(orderedFeatures[i].AssignedTo);
            }
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "O"]];
            Utility.AddNativieResource(range);
            range.Value2 = arr;

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);

            return nextRow;

        }
    }
}
