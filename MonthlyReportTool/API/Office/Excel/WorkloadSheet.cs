using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.TeamProject;
using MonthlyReportTool.API.TFS.WorkItem;
using MonthlyReportTool.API.TFS.Agile;

namespace MonthlyReportTool.API.Office.Excel
{
    public class WorkloadSheet : ExcelSheetBase, IExcelSheet
    {
        private ExcelInterop.Worksheet sheet;
        private ProjectEntity project;
        private List<WorkloadEntity> workloadList;
        public WorkloadSheet(ExcelInterop.Worksheet sheet) : base(sheet)
        {
            this.sheet = sheet;
        }

        public void Build(ProjectEntity project)
        {
            this.project = project;
            this.workloadList = TFS.WorkItem.Workload.GetAll(project.Name, TFS.Utility.GetBestIteration(project.Name));
            BuildTitle();
            BuildSubTitle();
            BuildDescription();
            BuildSummaryTable();

            BuildDevelopmentTitle();
            BuildDevelopmentDescription();

            int startRow = BuildDevelopmentTable();
            int dataRow = startRow - 3;

            List<Tuple<string, double, double>> workloads = new List<Tuple<string, double, double>>();
            for (int i = 14; i <= dataRow; i++)
            {
                string textb = this.sheet.Cells[i, "B"].Text;
                string textf = this.sheet.Cells[i, "F"].Text;
                string texti = this.sheet.Cells[i, "I"].Text;

                if (texti.StartsWith("#")) texti = "0%";

                if (String.IsNullOrEmpty(textb)) continue;

                workloads.Add(Tuple.Create<string, double, double>(this.sheet.Cells[i, "B"].Text, Convert.ToDouble(textf.Replace("%", "")) / 100.00d, Convert.ToDouble(texti.Replace("%", "")) / 100.00d));
            }

            startRow = Build120Analysis(startRow, workloads);
            startRow = Build100Analysis(startRow, workloads);
            startRow = Build60Analysis(startRow, workloads);
            startRow = BUild10Analysis(startRow, dataRow);

            ExcelInterop.Range colb = sheet.Cells[1, "B"];
            Utility.AddNativieResource(colb);
            colb.ColumnWidth = 10;

            sheet.Cells[1, "A"] = "";

        }
        private void BuildTitle()
        {
            Utility.BuildFormalSheetTitle(sheet, 2, "B", 2, "AG", "工作量统计分析");
        }
        private void BuildSubTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[4, "B"], sheet.Cells[4, "O"]];
            Utility.AddNativieResource(range);
            range.Merge();
            sheet.Cells[4, "B"] = "团队整体工作量分析";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;
            titleFont.Size = 12;

        }
        private void BuildDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[5, "B"], sheet.Cells[5, "AG"]];
            Utility.AddNativieResource(titleRange);
            titleRange.Merge();
            sheet.Cells[5, "B"] = "说明：容量投入总工时：迭代成员迭代规划容量×迭代天数之和计算";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;
        }
        private int BuildSummaryTable()
        {
            int featuresCount = 1;
            string[] cols = new string[] { "团队成员标准工时\r\n（迭代天数*人数×8）", "实际投入总工时", "工作\r\n饱和度", "研发投入总工时", "研发投入\r\n占比", "容量投入总工时\r\n（迭代天数×人数×容量）", "评估总工时\r\n（迭代任务的评估工时）", "迭代任务的实际工时", "计划偏差", "剩余工时" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","E"),
                Tuple.Create<string,string>("F","H"),
                Tuple.Create<string,string>("I","J"),
                Tuple.Create<string,string>("K","L"),
                Tuple.Create<string,string>("M","N"),
                Tuple.Create<string,string>("O","Q"),
                Tuple.Create<string,string>("R","T"),
                Tuple.Create<string,string>("U","W"),
                Tuple.Create<string,string>("X","Z"),
                Tuple.Create<string,string>("AA","AC"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[6, colsname[i].Item1], sheet.Cells[6, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.Merge();
                sheet.Cells[6, colsname[i].Item1] = cols[i];
            }

            ExcelInterop.Range firstRow = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "AC"]];
            Utility.AddNativieResource(firstRow);
            var border = firstRow.Borders;
            Utility.AddNativieResource(border);
            border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            firstRow.Copy();

            ExcelInterop.Range nextRow = sheet.Range[sheet.Cells[7, "B"], sheet.Cells[7, "AC"]];
            Utility.AddNativieResource(nextRow);
            nextRow.PasteSpecial(ExcelInterop.XlPasteType.xlPasteFormats);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[7, "B"], sheet.Cells[7, "N"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[7, "X"], sheet.Cells[7, "Z"]]);

            firstRow.RowHeight = 50;
            firstRow.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = firstRow.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.LightGray.ToArgb();

            return 7 + featuresCount + 2;
        }

        private void BuildDevelopmentTitle()
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[9, "B"], sheet.Cells[9, "AG"]];
            Utility.AddNativieResource(range);
            range.RowHeight = 20;
            range.Merge();
            sheet.Cells[9, "B"] = "开发工作量统计";
            var titleFont = range.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Bold = true;

            ExcelInterop.Range colA = sheet.Cells[1, "A"] as ExcelInterop.Range;
            Utility.AddNativieResource(colA);
            colA.ColumnWidth = 2;
        }

        private void BuildDevelopmentDescription()
        {
            ExcelInterop.Range titleRange = sheet.Range[sheet.Cells[10, "B"], sheet.Cells[10, "K"]];
            Utility.AddNativieResource(titleRange);
            titleRange.RowHeight = 120;
            titleRange.Merge();
            sheet.Cells[10, "B"] = "说明：包括测试人员的工作量统计\r\n" +
           "        按照实际饱和度排序（开发人员和测试人员分开排序）\r\n" +
           "        各明细工作量是实际填写的工作日志某类的工作量，和工作日志中分类一致\r\n" +
           "        Bug数：对于开发人员是本迭代被测试出的Bug总数\r\n" +
           "                 对于测试人员是本迭代测试出的Bug总数";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Size = 11;

            Utility.SetCellColor(sheet.Cells[10, "B"], System.Drawing.Color.Red, "按照实际饱和度排序（开发人员和测试人员分开排序）");

            ExcelInterop.Range titleRange2 = sheet.Range[sheet.Cells[10, "L"], sheet.Cells[10, "W"]];
            titleRange2.Merge();
            Utility.AddNativieResource(titleRange2);
            titleRange.Merge();
            sheet.Cells[10, "L"] = "      标准工作量：迭代天数*8（跨多个团队项目的人员，请手工调整每个团队投入的标准工作量=迭代天数×计划投入到此团队的占的8里的数据）\r\n" +
                                   "      实际投入工作量：不包含请假实际的工作量\r\n" +
                                   "      实际饱和度：实际投入工作量 / 标准工作量\r\n" +
                                   "      Bug产出率：bug数 / 实际投入工作量";
            Utility.SetCellColor(sheet.Cells[10, "L"], System.Drawing.Color.Red, "跨多个团队项目的人员，请手工调整每个团队投入的标准工作量=迭代天数×计划投入到此团队的占的8里的数据", true);
        }

        private int BUild10Analysis(int startRow, int dataRow)
        {

            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "计划偏差大于10%分析", "说明：", "B", "U",
                new List<string>() { "计划偏差大于10%分析", "说明：", "计划偏差大于10%分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                5,
                nodata: true
                );

            sheet.Cells[startRow + 2, "B"] = "计划偏差大于10%分析";
            Utility.SetCellFontRedColor(sheet.Cells[startRow + 2, "B"]);

            //ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2 + 5, "U"]];
            //Utility.AddNativieResource(range);
            //range.Merge();

            return nextRow - 1;
        }

        private int Build60Analysis(int startRow, List<Tuple<string, double, double>> workloads)
        {
            var ds = workloads.Where(wl => wl.Item3 <= 0.60d).OrderByDescending(wl => wl.Item3).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "研发占比低于60%分析", "说明：对研发占比低于60%（不包括60%）的同事，分别做原因分析", "B", "U",
                new List<string>() { "团队成员", "研发占比", "研发占比低于60%原因分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                ds.Count()
                );

            Utility.SetCellFontRedColor(sheet.Cells[startRow + 2, "F"]);
            startRow += 3;
            for (int i = 0; i < ds.Count(); i++)
            {
                this.sheet.Cells[startRow + i, "B"] = ds[i].Item1;
                this.sheet.Cells[startRow + i, "C"] = ds[i].Item3;
                this.sheet.Cells[startRow + i, "F"] = "";
            }

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + ds.Count - 1, "B"]]);

            Utility.SetCellPercentFormat(sheet.Range[sheet.Cells[startRow, "C"], sheet.Cells[startRow + ds.Count() - 1, "C"]]);

            return nextRow - 1;
        }

        private int Build120Analysis(int startRow, List<Tuple<string, double, double>> workloads)
        {
            var ds = workloads.Where(wl => wl.Item2 >= 1.20d).OrderByDescending(wl => wl.Item2).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "工作量饱和度超120%分析", "说明：对工作量饱和度超过120%（包括120%）的同事，分别做原因分析", "B", "U",
                new List<string>() { "团队成员", "工作量饱和度", "工作量饱和度超120%原因分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                ds.Count()
                );

            Utility.SetCellFontRedColor(sheet.Cells[startRow + 2, "F"]);
            startRow += 3;
            for (int i = 0; i < ds.Count(); i++)
            {
                this.sheet.Cells[startRow + i, "B"] = ds[i].Item1;
                this.sheet.Cells[startRow + i, "C"] = ds[i].Item2;
                this.sheet.Cells[startRow + i, "F"] = "";
            }

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + ds.Count - 1, "B"]]);

            Utility.SetCellPercentFormat(sheet.Range[sheet.Cells[startRow, "C"], sheet.Cells[startRow + ds.Count() - 1, "C"]]);

            return nextRow - 1;
        }

        private int Build100Analysis(int startRow, List<Tuple<string, double, double>> workloads)
        {
            var ds = workloads.Where(wl => wl.Item2 < 1.00d).OrderByDescending(wl => wl.Item2).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "工作量饱和度不足100%分析", "说明：对工作量饱和度不足100%%（不包括100%）的同事，分别做原因分析", "B", "U",
                new List<string>() { "团队成员", "工作量饱和度", "工作量饱和度不足100%原因分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                ds.Count()
                );

            Utility.SetCellFontRedColor(sheet.Cells[startRow + 2, "F"]);
            startRow += 3;
            for (int i = 0; i < ds.Count(); i++)
            {
                this.sheet.Cells[startRow + i, "B"] = ds[i].Item1;
                this.sheet.Cells[startRow + i, "C"] = ds[i].Item2;
                this.sheet.Cells[startRow + i, "F"] = "";
            }

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow + ds.Count - 1, "B"]]);

            Utility.SetCellPercentFormat(sheet.Range[sheet.Cells[startRow, "C"], sheet.Cells[startRow + ds.Count() - 1, "C"]]);

            return nextRow - 1;
        }

        private int FillWorkloadData(List<WorkloadEntity> workloads, int startrow, bool isTester)
        {
            int firstrow = startrow;

            int standardWorkingDays = TFS.Utility.GetStandardWorkingDays(this.project.Name, TFS.Utility.GetBestIteration(this.project.Name));
            var orderedLoads = GetOrderedWorkloads(workloads);

            var allbugs = Bug.GetAddedBugsByIteration(this.project.Name, API.TFS.Utility.GetBestIteration(this.project.Name));
            object[,] arr = new object[orderedLoads.Count(), 33];
            int line = 0;

            var capacities = TFS.Agile.Capacity.GetIterationCapacitiesForTeamMember(this.project.Name, TFS.Utility.GetBestIteration(this.project.Name).Id);

            foreach (var load in orderedLoads)
            {
                string person = Utility.GetPersonName(load.Key);
                int bugCount = 0;
                if (isTester)
                {
                    bugCount = allbugs.Where(bug => bug.DiscoveryUser == load.Key).Count();
                }
                else
                {
                    bugCount = allbugs.Where(bug => bug.AssignedTo == load.Key).Count();
                }
                double leavetime = load.Where(wl => wl.Type == "请假").Sum(wl => wl.SumHours);

                var dev = load.Where(wl => wl.SupperType == "研发");
                double devtime = dev.Sum(wl => wl.SumHours);
                double devtime1 = dev.Where(wl => wl.Type == "开发").Sum(wl => wl.SumHours);
                double devtime2 = dev.Where(wl => wl.Type == "需求").Sum(wl => wl.SumHours);
                double devtime3 = dev.Where(wl => wl.Type == "设计").Sum(wl => wl.SumHours);
                double devtime4 = dev.Where(wl => wl.Type == "测试设计").Sum(wl => wl.SumHours);
                double devtime5 = dev.Where(wl => wl.Type == "测试执行").Sum(wl => wl.SumHours);
                double devtime6 = devtime - (devtime1 + devtime2 + devtime3 + devtime4 + devtime5);

                double mgrtime = load.Where(wl => wl.SupperType == "管理").Sum(wl => wl.SumHours);
                double oprtime = load.Where(wl => wl.SupperType == "运维").Sum(wl => wl.SumHours);
                double doctime = load.Where(wl => wl.SupperType == "文档").Sum(wl => wl.SumHours);
                double studytime = load.Where(wl => wl.SupperType == "学习交流").Sum(wl => wl.SumHours);
                double presalestime = load.Where(wl => wl.SupperType == "售前/推广").Sum(wl => wl.SumHours);

                double othertime = load.Where(wl => wl.SupperType == "其他").Sum(wl => wl.SumHours) - leavetime;

                arr[line, 0] = person;
                //if (capacities.ContainsKey(load.Key))
                //{
                //    arr[line, 1] = capacities[load.Key] * standardWorkingDays;
                //}
                //else
                //{
                //    arr[line, 1] = 9999;// "未在迭代settings中定义【"+Utility.GetPersonName(load.Key)+"】的容量";
                //}
                arr[line, 1] = standardWorkingDays * 8;//capacities[load.Key] * standardWorkingDays;// standardWorkingDays * 8;
                arr[line, 2] = load.Where(wl => wl.Type != "请假").Sum(wl => wl.SumHours);
                arr[line, 3] = leavetime;
                arr[line, 4] = String.Format("=D{0}/C{0}", startrow);
                arr[line, 5] = bugCount;
                arr[line, 6] = String.Format("=G{0}/(D{0}/8)", startrow);
                arr[line, 7] = devtime / (load.Sum(wl => wl.SumHours) - leavetime);
                arr[line, 8] = devtime1;
                arr[line, 9] = String.Format("=IF(J{0}<>0,J{0}/D{0}, \"\"", startrow);
                arr[line, 10] = devtime2;
                arr[line, 11] = String.Format("=IF(L{0}<>0,L{0}/D{0}, \"\"", startrow);
                arr[line, 12] = devtime3;
                arr[line, 13] = String.Format("=IF(N{0}<>0,N{0}/D{0}, \"\"", startrow);
                arr[line, 14] = devtime4;
                arr[line, 15] = String.Format("=IF(P{0}<>0,P{0}/D{0}, \"\"", startrow);
                arr[line, 16] = devtime5;
                arr[line, 17] = String.Format("=IF(R{0}<>0,R{0}/D{0}, \"\"", startrow);
                arr[line, 18] = devtime6;
                arr[line, 19] = String.Format("=IF(T{0}<>0,T{0}/D{0}, \"\"", startrow);

                arr[line, 20] = mgrtime;
                arr[line, 21] = String.Format("=IF(V{0}<>0,V{0}/D{0}, \"\"", startrow);
                arr[line, 22] = oprtime;
                arr[line, 23] = String.Format("=IF(X{0}<>0,X{0}/D{0}, \"\"", startrow);
                arr[line, 24] = doctime;
                arr[line, 25] = String.Format("=IF(Z{0}<>0,Z{0}/D{0}, \"\"", startrow);
                arr[line, 26] = studytime;
                arr[line, 27] = String.Format("=IF(AB{0}<>0,AB{0}/D{0}, \"\"", startrow);
                arr[line, 28] = presalestime;
                arr[line, 29] = String.Format("=IF(AD{0}<>0,AD{0}/D{0}, \"\"", startrow);
                arr[line, 30] = othertime;
                arr[line, 31] = String.Format("=IF(AF{0}<>0,AF{0}/D{0}, \"\"", startrow);

                startrow++;
                line++;
            }

            ExcelInterop.Range devRange = sheet.Range[sheet.Cells[firstrow, "B"], sheet.Cells[startrow - 1, "AG"]];
            Utility.AddNativieResource(devRange);
            devRange.Value2 = arr;
            devRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignRight;
            devRange.WrapText = true;
            devRange.RowHeight = 20;
            devRange.ColumnWidth = 7;

            Utility.SetCellBorder(devRange);

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[firstrow, "B"], sheet.Cells[firstrow + orderedLoads.Count() - 1, "B"]]);

            Utility.SetCellPercentFormat(sheet.Range[sheet.Cells[firstrow, "F"], sheet.Cells[startrow - 1, "F"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "D"], sheet.Cells[startrow - 1, "D"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "F"], sheet.Cells[startrow - 1, "F"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "H"], sheet.Cells[startrow - 1, "H"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "I"], sheet.Cells[startrow - 1, "I"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "K"], sheet.Cells[startrow - 1, "K"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "M"], sheet.Cells[startrow - 1, "M"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "O"], sheet.Cells[startrow - 1, "O"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "Q"], sheet.Cells[startrow - 1, "Q"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "S"], sheet.Cells[startrow - 1, "S"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "U"], sheet.Cells[startrow - 1, "U"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "W"], sheet.Cells[startrow - 1, "W"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "Y"], sheet.Cells[startrow - 1, "Y"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "AA"], sheet.Cells[startrow - 1, "AA"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "AC"], sheet.Cells[startrow - 1, "AC"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "AE"], sheet.Cells[startrow - 1, "AE"]]);
            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[firstrow, "AG"], sheet.Cells[startrow - 1, "AG"]]);

            ExcelInterop.Range bugrange = sheet.Range[sheet.Cells[firstrow, "H"], sheet.Cells[startrow - 1, "H"]];
            Utility.AddNativieResource(bugrange);
            bugrange.NumberFormat = "#0.00";

            for (int col = 9; col <= 33; col += 2)
            {
                Utility.SetCellPercentFormat(sheet, firstrow, col, startrow - 1, col);
            }

            Utility.SetFormatBigger(sheet.Range[sheet.Cells[firstrow, "F"], sheet.Cells[startrow - 1, "F"]], 1.20d);
            Utility.SetFormatSmaller(sheet.Range[sheet.Cells[firstrow, "F"], sheet.Cells[startrow - 1, "F"]], 1.00d);
            Utility.SetFormatSmaller(sheet.Range[sheet.Cells[firstrow, "I"], sheet.Cells[startrow - 1, "I"]], 0.60d);

            return startrow;
        }

        private IOrderedEnumerable<IGrouping<string, WorkloadEntity>> GetOrderedWorkloads(List<WorkloadEntity> workloads)
        {
            var loads = workloads.GroupBy(wl => wl.AssignedTo);

            return loads.OrderByDescending(
                load => (
                    load.Where(eachload => eachload.Type != "请假").Sum(eachload => eachload.SumHours)
                    )
                );
        }

        private int BuildDevelopmentTable()
        {
            string[,] cols = new string[,]
                        {
                { "团队成员", "标准\r\n工作量", "实际投入\r\n工作量", "请假","实际\r\n饱和度","bug数","bug产出率（个/天）","研发工作量占比","研发","","","","","","","","","","","","管理","","运维","","文档","","学习交流","","售前/推广","","其他\r\n排除请假",""},
                { "", "", "", "", "", "", "", "", "开发","","需求","","设计","","测试设计","","测试执行","","其他","","","","","","","","","","","","",""},
                { "", "", "", "", "", "", "", "", "工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比","工作量","占比"},
                        };

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < cols.GetLength(1); col++)
                {
                    sheet.Cells[row + 11, col + 2] = cols[row, col];

                }
            }

            List<Tuple<int, string, int, string>> allMergedCells = new List<Tuple<int, string, int, string>>()
            {
                Tuple.Create<int, string, int, string>(11,"B",13,"B"),Tuple.Create<int, string, int, string>(11,"C",13,"C"),
                Tuple.Create<int, string, int, string>(11,"D",13,"D"),Tuple.Create<int, string, int, string>(11,"E",13,"E"),
                Tuple.Create<int, string, int, string>(11,"F",13,"F"),Tuple.Create<int, string, int, string>(11,"G",13,"G"),
                Tuple.Create<int, string, int, string>(11,"H",13,"H"),Tuple.Create<int, string, int, string>(11,"I",13,"I"),

                Tuple.Create<int, string, int, string>(11,"J",11,"U"),

                Tuple.Create<int, string, int, string>(11,"V",12,"W"),Tuple.Create<int, string, int, string>(11,"X",12,"Y"),
                Tuple.Create<int, string, int, string>(11,"Z",12,"AA"),Tuple.Create<int, string, int, string>(11,"AB",12,"AC"),
                Tuple.Create<int, string, int, string>(11,"AD",12,"AE"),Tuple.Create<int, string, int, string>(11,"AF",12,"AG"),

                Tuple.Create<int, string, int, string>(12,"J",12,"K"),Tuple.Create<int, string, int, string>(12,"L",12,"M"),
                Tuple.Create<int, string, int, string>(12,"N",12,"O"),Tuple.Create<int, string, int, string>(12,"P",12,"Q"),
                Tuple.Create<int, string, int, string>(12,"R",12,"S"),Tuple.Create<int, string, int, string>(12,"T",12,"U")
            };

            foreach (var tuple in allMergedCells)
            {
                ExcelInterop.Range range = sheet.Range[sheet.Cells[tuple.Item1, tuple.Item2], sheet.Cells[tuple.Item3, tuple.Item4]];
                Utility.AddNativieResource(range);
                range.Merge();
            }

            Utility.SetTableHeaderFormat(sheet.get_Range("B11:AG13"), false);

            var testlist = TFS.Utility.GetTestMembers(false);
            List<WorkloadEntity> devload = new List<WorkloadEntity>();
            List<WorkloadEntity> testload = new List<WorkloadEntity>();

            for (int i = 0; i < this.workloadList.Count(); i++)
            {
                bool isTester = testlist.Where(test => test == this.workloadList[i].AssignedTo).Count() > 0;
                if (isTester)
                {
                    testload.Add(this.workloadList[i]);
                }
                else
                {
                    devload.Add(this.workloadList[i]);
                }
            }

            int startrow = 14;
            startrow = FillWorkloadData(devload, startrow, false);
            startrow = FillWorkloadData(testload, startrow + 1, true);
            FillSummaryData(startrow);

            Utility.SetCellGreenColor(sheet.Range[sheet.Cells[startrow, "C"], sheet.Cells[startrow, "AG"]]);
            Utility.SetCellDarkGrayColor(sheet.Range[sheet.Cells[startrow, "B"], sheet.Cells[startrow, "B"]]);

            var ite = TFS.Utility.GetBestIteration(this.project.Name);
            //int totalDays = (DateTime.Parse(ite.EndDate).AddDays(1) - DateTime.Parse(ite.StartDate)).Days;
            //sheet.Cells[7, "O"] = TFS.Agile.Capacity.GetIterationCapacities(this.project.Name, ite.Id) * totalDays;
            int standardWorkingDays = TFS.Utility.GetStandardWorkingDays(this.project.Name, TFS.Utility.GetBestIteration(this.project.Name));
            sheet.Cells[7, "O"] = TFS.Agile.Capacity.GetIterationCapacities(this.project.Name, ite.Id) * standardWorkingDays;

            var estimated = TFS.WorkItem.Workload.GetEstimated(this.project.Name, ite);
            sheet.Cells[7, "R"] = estimated.Item1;
            sheet.Cells[7, "U"] = estimated.Item2;
            sheet.Cells[7, "X"] = "=(U7-R7)/R7";
            sheet.Cells[7, "AA"] = estimated.Item3;

            Utility.SetCellPercentFormat(sheet.Cells[7, "X"]);

            return startrow + 2;
        }

        private void FillSummaryData(int startRow)
        {
            Utility.SetCellBorder(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "AG"]]);

            sheet.Cells[startRow, "B"] = "合计";
            sheet.Cells[startRow, "C"] = String.Format("=sum(C14:C{0}", startRow - 1);
            sheet.Cells[startRow, "D"] = String.Format("=sum(D14:D{0}", startRow - 1);
            sheet.Cells[startRow, "E"] = String.Format("=sum(E14:E{0}", startRow - 1);
            sheet.Cells[startRow, "F"] = String.Format("=IF(C{0}<>0,D{0}/C{0},\"\")", startRow);
            sheet.Cells[startRow, "G"] = String.Format("=sum(G14:G{0}", startRow - 1);
            sheet.Cells[startRow, "H"] = String.Format("=G{0}/(D{0}/8)", startRow);
            sheet.Cells[startRow, "I"] = String.Format("=(J{0} + L{0} + N{0} + P{0} + R{0} + T{0}) / D{0}", startRow);
            sheet.Cells[startRow, "J"] = String.Format("=sum(J14:J{0}", startRow - 1);
            sheet.Cells[startRow, "K"] = String.Format("=J{0}/D{0}", startRow);
            sheet.Cells[startRow, "L"] = String.Format("=sum(L14:L{0}", startRow - 1);
            sheet.Cells[startRow, "M"] = String.Format("=L{0}/D{0}", startRow);
            sheet.Cells[startRow, "N"] = String.Format("=sum(N14:N{0}", startRow - 1);
            sheet.Cells[startRow, "O"] = String.Format("=N{0}/D{0}", startRow);
            sheet.Cells[startRow, "P"] = String.Format("=sum(P14:P{0}", startRow - 1);
            sheet.Cells[startRow, "Q"] = String.Format("=P{0}/D{0}", startRow);
            sheet.Cells[startRow, "R"] = String.Format("=sum(R14:R{0}", startRow - 1);
            sheet.Cells[startRow, "S"] = String.Format("=R{0}/D{0}", startRow);
            sheet.Cells[startRow, "T"] = String.Format("=sum(T14:T{0}", startRow - 1);
            sheet.Cells[startRow, "U"] = String.Format("=T{0}/D{0}", startRow);
            sheet.Cells[startRow, "V"] = String.Format("=sum(V14:V{0}", startRow - 1);
            sheet.Cells[startRow, "W"] = String.Format("=V{0}/D{0}", startRow);
            sheet.Cells[startRow, "X"] = String.Format("=sum(X14:X{0}", startRow - 1);
            sheet.Cells[startRow, "Y"] = String.Format("=X{0}/D{0}", startRow);
            sheet.Cells[startRow, "Z"] = String.Format("=sum(Z14:Z{0}", startRow - 1);
            sheet.Cells[startRow, "AA"] = String.Format("=Z{0}/D{0}", startRow);
            sheet.Cells[startRow, "AB"] = String.Format("=sum(AB14:AB{0}", startRow - 1);
            sheet.Cells[startRow, "AC"] = String.Format("=AB{0}/D{0}", startRow);
            sheet.Cells[startRow, "AD"] = String.Format("=sum(AD14:AD{0}", startRow - 1);
            sheet.Cells[startRow, "AE"] = String.Format("=AD{0}/D{0}", startRow);
            sheet.Cells[startRow, "AF"] = String.Format("=sum(AF14:AF{0}", startRow - 1);
            sheet.Cells[startRow, "AG"] = String.Format("=AF{0}/D{0}", startRow);

            sheet.Cells[7, "B"] = String.Format("=C{0}", startRow);
            sheet.Cells[7, "F"] = String.Format("=D{0}", startRow);
            sheet.Cells[7, "I"] = String.Format("=F7/B7", startRow);
            sheet.Cells[7, "K"] = String.Format("=J{0}+L{0}+N{0}+P{0}+R{0}+T{0}", startRow);
            sheet.Cells[7, "M"] = String.Format("=K7/F7", startRow);

            Utility.SetCellPercentFormat(sheet.Cells[startRow, "F"]);

            for (int i = 9; i <= 33; i += 2)
            {
                Utility.SetCellPercentFormat(sheet, startRow, i, startRow, i);
            }

            Utility.SetCellPercentFormat(sheet, 7, 9, 7, 9);
            Utility.SetCellPercentFormat(sheet, 7, 13, 7, 13);

            ExcelInterop.Range bugrange = sheet.Cells[startRow, "H"];
            Utility.AddNativieResource(bugrange);
            bugrange.NumberFormat = "#0.00";

            Utility.SetCellAlignAndWrap(sheet.Range[sheet.Cells[startRow, "B"], sheet.Cells[startRow, "B"]], hAlign: ExcelInterop.XlHAlign.xlHAlignCenter);
        }
    }
}
