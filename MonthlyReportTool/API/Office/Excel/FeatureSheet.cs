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
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "R", "需求统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "K"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "本迭代需求完成情况统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Name = "微软雅黑";
            titleFont.Size = 12;
        }
        private void BuildDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "K"]];
            Utility.AddNativieResource(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明：统计范围：目标日期在本迭代期间内的所有需求；";

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
            int nextRow = Utility.BuildFormalTable(sheet, startRow, "本迭代需求完成情况统计", "说明：统计范围：目标日期在本迭代期间内的所有需求；", "B", "K",
                new List<string>() { "分类", "个数", "占比", "说明" },
                new List<string>() { "B,D", "E,E", "F,F", "G,K" },
                5);

            string[] cols1 = new string[] { "已完成数", "未完成数", "中止/移除数", "按计划完成数", "应完成总数" };
            string[] cols2 = new string[] { "=IF(E11<>0,E7/E11,\"\")", "=IF(E11<>0,E8/E11,\"\")", "'--", "=IF(E11<>0,E10/E11,\"\")", "'--" };
            string[] cols3 = new string[] { "已完成数：目标日期内已发布的需求总数\r\n占比：已完成数/目标日期内应完成总数",
                "未完成数：目标日期内未完成的需求总数\r\n占比：未完成数/目标日期内应完成总数",
                "中止/移除数：目标日期内中止或移除的需求总数\r\n占比：中止/移除数/目标日期内应完成总数",
                "按计划完成数：按目标日期完成发布的需求总数\r\n占比：按计划完成数/目标日期内应完成总数",
                "所有应发布需求总数" };

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
            Utility.SetCellGreenColor(sheet.get_Range("F7:F11"));

            ExcelInterop.Range range = sheet.Range[sheet.Cells[7, "B"], sheet.Cells[12, "B"]];
            Utility.AddNativieResource(range);
            range.RowHeight = 40;

            Utility.SetFormatBigger(sheet.Cells[8, "E"], 0.0001d);
            Utility.SetFormatBigger(sheet.Cells[9, "E"], 0.0001d);
        }
        private int BuildTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代需求列表", "说明：如果一个单元格的内容太多，请考虑换行显示\r\n      按关键应用、模块、功能排序；", "B", "Q",
                new List<string>() { "ID", "关键应用", "模块", "功能", "需求描述", "是否需要需求分析", "当前状态", "目标日期", "已发布日期", "负责人" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,L", "M,M", "N,N", "O,O", "P,P", "Q,Q" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[14, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");

            var orderedFeatures = features.OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();
            startRow += 3;
            object[,] arr = new object[orderedFeatures.Count, 15];
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = orderedFeatures[i].Id;
                sheet.Cells[startRow + i, "C"] = orderedFeatures[i].KeyApplicationName;
                sheet.Cells[startRow + i, "E"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "G"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "I"] = orderedFeatures[i].Title;
                sheet.Cells[startRow + i, "M"] = orderedFeatures[i].NeedRequireDevelop;
                sheet.Cells[startRow + i, "N"] = orderedFeatures[i].State;
                sheet.Cells[startRow + i, "O"] = DateTime.Parse(orderedFeatures[i].TargetDate).AddHours(8).ToString("yyyy-MM-dd");
                if (String.IsNullOrEmpty(orderedFeatures[i].ReleaseFinishedDate))
                {
                    sheet.Cells[startRow + i, "P"] = "";
                }
                else
                {
                    sheet.Cells[startRow + i, "P"] = DateTime.Parse(orderedFeatures[i].ReleaseFinishedDate).AddHours(8).ToString("yyyy-MM-dd");
                }                

                sheet.Cells[startRow + i, "Q"] = Utility.GetPersonName(orderedFeatures[i].AssignedTo);
            }

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);
            return nextRow - 1;

            //ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "O"]];
            //Utility.AddNativieResource(range);
            //range.Value2 = arr;

            //Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);

            //return nextRow - 1;

        }

        private int BuildDelayTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代未完成需求分析", "说明：按关键应用、模块、功能排序；这个表格很长，请右拉把后面列都填写上。", "B", "R",
                new List<string>() { "ID", "关键应用", "模块", "功能", "需求描述", "是否需要需求分析", "当前状态", "目标日期", "负责人", "拖期原因说明" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,L", "M,M", "N,N", "O,O", "P,P", "Q,R" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");
            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "这个表格很长，请右拉把后面列都填写上。");
            var orderedFeatures = features.OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();
            startRow += 3;
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = orderedFeatures[i].Id;
                sheet.Cells[startRow + i, "C"] = orderedFeatures[i].KeyApplicationName;
                sheet.Cells[startRow + i, "E"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "G"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "I"] = orderedFeatures[i].Title;
                sheet.Cells[startRow + i, "M"] = orderedFeatures[i].NeedRequireDevelop;
                sheet.Cells[startRow + i, "N"] = orderedFeatures[i].State;
                sheet.Cells[startRow + i, "O"] = DateTime.Parse(orderedFeatures[i].TargetDate).AddHours(8).ToString("yyyy-MM-dd");
                //if (String.IsNullOrEmpty(orderedFeatures[i].ReleaseFinishedDate))
                //{
                //    sheet.Cells[startRow + i, "O"] = "";
                //}
                //else
                //{
                //    sheet.Cells[startRow + i, "O"] = DateTime.Parse(orderedFeatures[i].ReleaseFinishedDate).AddHours(8).ToString("yyyy-MM-dd");
                //}
                
                sheet.Cells[startRow + i, "P"] = Utility.GetPersonName(orderedFeatures[i].AssignedTo);
                sheet.Cells[startRow + i, "Q"] = "";
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "Q"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);

            return nextRow - 1;
        }
        private int BuildAnadonTable(int startRow, List<FeatureEntity> features)
        {
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "本迭代移除/中止需求分析", "说明：按关键应用、模块、功能排序；这个表格很长，请右拉把后面列都填写上。", "B", "R",
                new List<string>() { "ID", "关键应用", "模块", "功能", "需求描述", "是否需要需求分析", "当前状态", "目标日期", "负责人", "移除/中止原因说明" },
                new List<string>() { "B,B", "C,D", "E,F", "G,H", "I,L", "M,M", "N,N", "O,O", "P,P", "Q,R" },
                features.Count);

            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "按关键应用、模块、功能排序");
            Utility.SetCellColor(sheet.Cells[startRow + 1, "B"], System.Drawing.Color.Red, "这个表格很长，请右拉把后面列都填写上。");

            var orderedFeatures = features.OrderBy(feature => feature.KeyApplicationName).ThenBy(feature => feature.ModulesName).ThenBy(feature => feature.FuncName).ToList();
            startRow += 3;
            for (int i = 0; i < orderedFeatures.Count; i++)
            {
                sheet.Cells[startRow + i, "B"] = orderedFeatures[i].Id;
                sheet.Cells[startRow + i, "C"] = orderedFeatures[i].KeyApplicationName;
                sheet.Cells[startRow + i, "E"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "G"] = orderedFeatures[i].ModulesName;
                sheet.Cells[startRow + i, "I"] = orderedFeatures[i].Title;
                sheet.Cells[startRow + i, "M"] = orderedFeatures[i].NeedRequireDevelop;
                sheet.Cells[startRow + i, "N"] = orderedFeatures[i].State;
                sheet.Cells[startRow + i, "O"] = DateTime.Parse(orderedFeatures[i].TargetDate).AddHours(8).ToString("yyyy-MM-dd");
                                
                sheet.Cells[startRow + i, "P"] = Utility.GetPersonName(orderedFeatures[i].AssignedTo);
                sheet.Cells[startRow + i, "Q"] = "";
            }

            Utility.SetCellFontRedColor(sheet.Cells[startRow - 1, "Q"]);
            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + orderedFeatures.Count - 1, "B"]]);
            return nextRow - 1;
        }
    }
}
