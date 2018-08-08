using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.TeamProject;
using System.Collections;

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

            startRow = BuildDelayTable(startRow, this.featureList[3]);
            startRow = BuildAnadonTable(startRow, this.featureList[2]);

            var colKL = sheet.get_Range("B1:W1");
            Utility.AddNativieResource(colKL);
            colKL.ColumnWidth = 12;//6.27d;

            sheet.Cells[1, "A"] = "";
        }

        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "H", "系统需求统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "H"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "本迭代系统需求完成情况统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 12;
        }
        private void BuildDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "H"]];
            Utility.AddNativieResource(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明：统计范围：目标日期在本迭代期间内的所有系统需求；";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Name = "微软雅黑";
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            //tmpfont.Bold = true;
        }
        private void BuildSummaryTable(int startRow)
        {
            int nextRow = Utility.BuildFormalTable(sheet, startRow, "本迭代系统需求完成情况统计", "说明：统计范围：目标日期在本迭代期间内的所有明细系统需求（不需要需求分析的产品特性+需求分析工作项）；", "B", "I",
                new List<string>() { "分类", "个数", "占比", "说明" },
                new List<string>() { "B,B", "C,C", "D,D", "E,I" },
                5);

            string[] cols1 = new string[] { "已完成数", "未完成数", "中止/移除数", "按计划完成数", "应完成总数" };
            string[] cols2 = new string[] { "=IF(C11<>0,C7/C11,\"\")", "=IF(C11<>0,C8/C11,\"\")", "'--", "=IF(C11<>0,C10/C11,\"\")", "'--" };
            string[] cols3 = new string[] { "已完成数：迭代期间内已发布的需求总数\r\n占比：已完成数/迭代期间内应完成总数",
                "未完成数：迭代期间内未完成的需求总数\r\n占比：未完成数/迭代期间内应完成总数",
                "中止/移除数：迭代期间内中止或移除的需求总数\r\n占比：中止/移除数/迭代期间内应完成总数",
                "按计划完成数：按目标日期完成发布的需求总数\r\n占比：按计划完成数/迭代期间内应完成总数",
                "迭代期间内所有应发布需求总数" };

            for (int row = 7; row <= 11; row++)
            {
                sheet.Cells[row, "B"] = cols1[row - 7];
                sheet.Cells[row, "D"] = cols2[row - 7];
                sheet.Cells[row, "E"] = cols3[row - 7];
            }

            sheet.Cells[7, "C"] = this.featureList[6].Count;
            sheet.Cells[8, "C"] = this.featureList[8].Count;
            sheet.Cells[9, "C"] = this.featureList[7].Count;
            sheet.Cells[10, "C"] = this.featureList[9].Count;
            sheet.Cells[11, "C"] = this.featureList[5].Count;
            //sheet.Cells[12, "E"] = "=SUM(E7: E11)";

            Utility.SetCellPercentFormat(sheet.get_Range("D7:D11"));
            Utility.SetCellGreenColor(sheet.get_Range("D7:D11"));

            ExcelInterop.Range range = sheet.Range[sheet.Cells[7, "B"], sheet.Cells[11, "B"]];
            Utility.AddNativieResource(range);
            range.RowHeight = 40;

            Utility.SetFormatBigger(sheet.Cells[8, "C"], 0.0001d);
            Utility.SetFormatBigger(sheet.Cells[9, "C"], 0.0001d);
        }
        private int BuildTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代所有系统需求列表", "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      按关键应用、模块、功能排序；", "B", "S",
                new List<string>() { "ID", "关键应用", "模块", "功能", "需求描述", "是否需要需求分析", "当前状态", "目标日期", "计划需求分析完成日期","实际需求分析完成日期","已发布日期", "负责人" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,L", "M,M", "N,N", "O,O", "P,P", "Q,Q","R,R","S,S" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[14, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");

            //var orderedFeatures = features.OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();
            var orderedFeatures = features.FindAll(featrue => featrue.ParentId == "").OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();

            startRow += 3;
            int writeRow = startRow;

            object[,] arr = new object[orderedFeatures.Count, 15];
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                var feature = orderedFeatures[i];
                //已经写入，则继续
                UpdateOneRowForAllFeatures(writeRow, feature, false);
                writeRow++;
                //处理下级
                var childFeatrues = features.FindAll(featrue1 => featrue1.ParentId == Convert.ToString(feature.Id)).OrderBy(feature1 => feature.ParentId).ThenBy(feature1 => feature1.KeyApplicationName).ThenBy(feature1 => feature1.ModulesName).ThenBy(feature1 => feature1.FuncName).ToList();
                foreach (var feature1 in childFeatrues)
                {
                    UpdateOneRowForAllFeatures(writeRow, feature1, true);
                    writeRow++;
                }
            }
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + features.Count - 1, "B"]]);
            return nextRow - 1;
        }
        //更新所有系统需求列表写入
        private void UpdateOneRowForAllFeatures(int currentRow, FeatureEntity feature, bool child)
        {
            sheet.Cells[currentRow, "B"] = (child == true) ? "'  " + feature.Id : "" + feature.Id;
            sheet.Cells[currentRow, "C"] = feature.KeyApplicationName;
            sheet.Cells[currentRow, "E"] = feature.ModulesName;
            sheet.Cells[currentRow, "G"] = feature.ModulesName;
            sheet.Cells[currentRow, "I"] = (child==true)?"  "+feature.Title:feature.Title;
            sheet.Cells[currentRow, "M"] = feature.NeedRequireDevelop;
            sheet.Cells[currentRow, "N"] = feature.State;
            sheet.Cells[currentRow, "O"] = DateTime.Parse(feature.TargetDate).AddHours(8).ToString("yyyy-MM-dd");
            if (String.IsNullOrEmpty(feature.PlanRequireFinishDate))
            {
                sheet.Cells[currentRow, "P"] = "";
            }
            else
            {
                sheet.Cells[currentRow, "P"] = DateTime.Parse(feature.PlanRequireFinishDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            if (String.IsNullOrEmpty(feature.RequireFinishedDate))
            {
                sheet.Cells[currentRow, "Q"] = "";
            }
            else
            {
                sheet.Cells[currentRow, "Q"] = DateTime.Parse(feature.RequireFinishedDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            if (String.IsNullOrEmpty(feature.ReleaseFinishedDate))
            {
                sheet.Cells[currentRow, "R"] = "";
            }
            else
            {
                sheet.Cells[currentRow, "R"] = DateTime.Parse(feature.ReleaseFinishedDate).AddHours(8).ToString("yyyy-MM-dd");
            }

            sheet.Cells[currentRow, "S"] = Utility.GetPersonName(feature.AssignedTo);
        }

        private int BuildDelayTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代未发布系统需求分析（目标日期在本迭代内，但是迭代结束还未发布的系统需求）", "说明：按关键应用、模块、功能排序；这个表格很长，请右拉把后面列都填写上。", "B", "W",
                new List<string>() { "ID", "关键应用", "模块", "功能", "需求描述", "是否需要需求分析", "当前状态", "目标日期", "计划需求分析完成日期", "实际需求分析完成日期", "已发布日期","负责人", "未完成原因分析" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,L", "M,M", "N,N", "O,O", "P,P","Q,Q","R,R", "S,S","T,W" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");
            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "这个表格很长，请右拉把后面列都填写上。");

            //var orderedFeatures = features.OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();
            var orderedFeatures = features.FindAll(featrue => featrue.ParentId == "").OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();

            startRow += 3;

            int writeRow = startRow;
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                var feature = orderedFeatures[i];
                //已经写入，则继续
                UpdateOneRowForDelayFeatures(writeRow, feature, false);
                writeRow++;
                //处理下级
                var childFeatrues = features.FindAll(featrue1 => featrue1.ParentId == Convert.ToString(feature.Id)).OrderBy(feature1 => feature.ParentId).ThenBy(feature1 => feature1.KeyApplicationName).ThenBy(feature1 => feature1.ModulesName).ThenBy(feature1 => feature1.FuncName).ToList();
                foreach (var feature1 in childFeatrues)
                {
                    UpdateOneRowForDelayFeatures(writeRow, feature1, true);
                    writeRow++;
                }
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "T"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + features.Count - 1, "B"]]);

            return nextRow - 1;
        }

        //更新所有拖期系统需求列表写入
        private void UpdateOneRowForDelayFeatures(int currentRow, FeatureEntity feature, bool child)
        {
            sheet.Cells[currentRow, "B"] = (child == true) ? "'  " + feature.Id : "" + feature.Id;
            sheet.Cells[currentRow, "C"] = feature.KeyApplicationName;
            sheet.Cells[currentRow, "E"] = feature.ModulesName;
            sheet.Cells[currentRow, "G"] = feature.ModulesName;
            sheet.Cells[currentRow, "I"] = (child == true) ? "  " + feature.Title : feature.Title;
            sheet.Cells[currentRow, "M"] = feature.NeedRequireDevelop;
            sheet.Cells[currentRow, "N"] = feature.State;
            sheet.Cells[currentRow, "O"] = DateTime.Parse(feature.TargetDate).AddHours(8).ToString("yyyy-MM-dd");
            if (String.IsNullOrEmpty(feature.PlanRequireFinishDate)==false)
            {
                sheet.Cells[currentRow, "P"] = DateTime.Parse(feature.PlanRequireFinishDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            if (String.IsNullOrEmpty(feature.RequireFinishedDate)==false)
            {
                sheet.Cells[currentRow, "Q"] = DateTime.Parse(feature.RequireFinishedDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            if (String.IsNullOrEmpty(feature.ReleaseFinishedDate) == false)
            {
                sheet.Cells[currentRow, "R"] = DateTime.Parse(feature.ReleaseFinishedDate).AddHours(8).ToString("yyyy-MM-dd"); 
            }
            sheet.Cells[currentRow, "S"] = Utility.GetPersonName(feature.AssignedTo);
            sheet.Cells[currentRow, "T"] = "";
        }
        private int BuildAnadonTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代移除/中止系统需求分析", "说明：按关键应用、模块、功能排序；这个表格很长，请右拉把后面列都填写上。", "B", "W",
                new List<string>() { "ID", "关键应用", "模块", "功能", "需求描述", "是否需要需求分析", "当前状态", "目标日期", "计划需求分析完成日期", "实际需求分析完成日期","已移除/中止日期", "负责人", "移除/中止原因说明" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,L", "M,M", "N,N", "O,O", "P,P","Q,Q","R,R","S,S", "T,W" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");
            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "这个表格很长，请右拉把后面列都填写上。");

           // var orderedFeatures = features.OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();
            var orderedFeatures = features.FindAll(featrue => featrue.ParentId == "").OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();

            startRow += 3;
            int writeRow = startRow;
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                var feature = orderedFeatures[i];
                //已经写入，则继续
                UpdateOneRowForAnadonFeatures(writeRow, feature, false);
                writeRow++;
                //处理下级
                var childFeatrues = features.FindAll(featrue1 => featrue1.ParentId == Convert.ToString(feature.Id)).OrderBy(feature1 => feature.ParentId).ThenBy(feature1 => feature1.KeyApplicationName).ThenBy(feature1 => feature1.ModulesName).ThenBy(feature1 => feature1.FuncName).ToList();
                foreach (var feature1 in childFeatrues)
                {
                    UpdateOneRowForAnadonFeatures(writeRow, feature1, true);
                    writeRow++;
                }
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "T"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + features.Count - 1, "B"]]);
            return nextRow - 1;
        }

        //更新所有拖期系统需求列表写入
        private void UpdateOneRowForAnadonFeatures(int currentRow, FeatureEntity feature, bool child)
        {
            sheet.Cells[currentRow, "B"] = (child == true) ? "'  " + feature.Id : "" + feature.Id;
            sheet.Cells[currentRow, "C"] = feature.KeyApplicationName;
            sheet.Cells[currentRow, "E"] = feature.ModulesName;
            sheet.Cells[currentRow, "G"] = feature.ModulesName;
            sheet.Cells[currentRow, "I"] = (child == true) ? "  " + feature.Title : feature.Title;
            sheet.Cells[currentRow, "M"] = feature.NeedRequireDevelop;
            sheet.Cells[currentRow, "N"] = feature.State;
            sheet.Cells[currentRow, "O"] = DateTime.Parse(feature.TargetDate).AddHours(8).ToString("yyyy-MM-dd");
            if (String.IsNullOrEmpty(feature.PlanRequireFinishDate)==false)
            {
                sheet.Cells[currentRow, "P"] = DateTime.Parse(feature.PlanRequireFinishDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            if (String.IsNullOrEmpty(feature.RequireFinishedDate)==false)
            {
                sheet.Cells[currentRow, "Q"] = DateTime.Parse(feature.RequireFinishedDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            if (String.IsNullOrEmpty(feature.ClosedDate) == false)
            {
                sheet.Cells[currentRow, "R"] = DateTime.Parse(feature.ClosedDate).AddHours(8).ToString("yyyy-MM-dd");
            }
            sheet.Cells[currentRow, "S"] = Utility.GetPersonName(feature.AssignedTo);
            sheet.Cells[currentRow, "T"] = "";
        }
    }
}
