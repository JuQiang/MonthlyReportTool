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

            List<Tuple<string, double,double>> workloads = new List<Tuple<string, double,double>>();
            for (int i = 14; i <= dataRow; i++)
            {
                string textb = this.sheet.Cells[i, "B"].Text;
                string textf = this.sheet.Cells[i, "F"].Text;
                string texti = this.sheet.Cells[i, "I"].Text;

                if (String.IsNullOrEmpty(textb)) continue;

                workloads.Add(Tuple.Create<string, double,double>(this.sheet.Cells[i, "B"].Text, Convert.ToDouble(textf.Replace("%",""))/100.00d, Convert.ToDouble(texti.Replace("%", "")) / 100.00d));
            }

            startRow = Build115Analysis(startRow,workloads);
            startRow = Build60Analysis(startRow, workloads);
            startRow = BUild50Analysis(startRow, dataRow);

        }

        private int BUild50Analysis(int startRow, int dataRow)
        {
            
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "计划偏差大于50%分析", "说明：", "B", "U",
                new List<string>() { "计划偏差大于50%分析", "说明：", "研发占比低于60%原因分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                5
                );

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow + 2, "B"], sheet.Cells[startRow + 2 + 5, "U"]];
            Utility.AddNativieResource(range);
            range.Merge();

            return nextRow;
        }

        private int Build60Analysis(int startRow, List<Tuple<string, double, double>> workloads)
        {
            var ds = workloads.Where(wl => wl.Item3 <= 0.60d).OrderByDescending(wl=>wl.Item3).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "研发占比低于60%分析", "说明：对研发占比低于60%（不包括60%）的同事，分别做原因分析", "B", "U",
                new List<string>() { "团队成员", "研发占比", "研发占比低于60%原因分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                ds.Count()
                );

            for (int i = 0; i < ds.Count(); i++)
            {
                this.sheet.Cells[startRow + 3 + i, "B"] = ds[i].Item1;
                this.sheet.Cells[startRow + 3 + i, "C"] = ds[i].Item3;
                this.sheet.Cells[startRow + 3 + i, "F"] = "";
            }

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow + 3, "B"], sheet.Cells[startRow + 3 + ds.Count() - 1, "B"]];
            Utility.AddNativieResource(range);
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            ExcelInterop.Range range2 = sheet.Range[sheet.Cells[startRow + 3, "C"], sheet.Cells[startRow + 3 + ds.Count() - 1, "C"]];
            Utility.AddNativieResource(range2);
            range2.NumberFormat = "#%";

            return nextRow;
        }

        private int Build115Analysis(int startRow,List<Tuple<string, double, double>> workloads)
        {

            var ds = workloads.Where(wl => wl.Item2 >= 1.15d).ToList();
            int nextRow = Utility.BuildFormalTable(this.sheet, startRow, "工作量饱和度超115%分析", "说明：对工作量饱和度超过115%（包括115%）的同事，分别做原因分析", "B", "U",
                new List<string>() { "团队成员", "工作量饱和度", "工作量饱和度超115%原因分析" },
                new List<string>() { "B,B", "C,E", "F,U" },
                ds.Count()
                );

            for (int i = 0; i < ds.Count(); i++)
            {
                this.sheet.Cells[startRow + 3 + i, "B"] = ds[i].Item1;
                this.sheet.Cells[startRow + 3 + i, "C"] = ds[i].Item2;
                this.sheet.Cells[startRow + 3 + i, "F"] = "";
            }

            ExcelInterop.Range range = sheet.Range[sheet.Cells[startRow + 3, "B"], sheet.Cells[startRow + 3+ds.Count()-1, "B"]];
            Utility.AddNativieResource(range);
            range.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            ExcelInterop.Range range2 = sheet.Range[sheet.Cells[startRow + 3, "C"], sheet.Cells[startRow + 3 + ds.Count() - 1, "C"]];
            Utility.AddNativieResource(range2);
            range2.NumberFormat = "#%";

            return nextRow;
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
            string[] cols = new string[] { "团队成员标准工时\r\n（迭代天数*人数×8）", "实际投入总工时", "工作\r\n饱和度", "研发投入总工时", "研发投入\r\n占比", "容量投入总工时\r\n（迭代天数×人数×容量）", "评估总工时\r\n（迭代任务的评估工时）", "剩余工时" };
            List<Tuple<string, string>> colsname = new List<Tuple<string, string>>(){
                Tuple.Create<string,string>("B","E"),
                Tuple.Create<string,string>("F","H"),
                Tuple.Create<string,string>("I","J"),
                Tuple.Create<string,string>("K","L"),
                Tuple.Create<string,string>("M","N"),
                Tuple.Create<string,string>("O","Q"),
                Tuple.Create<string,string>("R","T"),
                Tuple.Create<string,string>("U","W"),
            };

            for (int i = 0; i < cols.Length; i++)
            {
                ExcelInterop.Range colRange = sheet.Range[sheet.Cells[6, colsname[i].Item1], sheet.Cells[6, colsname[i].Item2]];
                Utility.AddNativieResource(colRange);
                colRange.Merge();
                sheet.Cells[6, colsname[i].Item1] = cols[i];

                var border = colRange.Borders;
                Utility.AddNativieResource(border);
                border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }

            ExcelInterop.Range tableRange = sheet.Range[sheet.Cells[6, "B"], sheet.Cells[6, "W"]];
            Utility.AddNativieResource(tableRange);
            tableRange.RowHeight = 50;
            tableRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            var interior = tableRange.Interior;
            Utility.AddNativieResource(interior);
            interior.Color = System.Drawing.Color.DarkGray.ToArgb();

            //TODO : 放入GIT
            for (int i = 0; i < featuresCount; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    ExcelInterop.Range colRange = sheet.Range[sheet.Cells[7 + i, colsname[j].Item1], sheet.Cells[7 + i, colsname[j].Item2]];
                    Utility.AddNativieResource(colRange);
                    colRange.RowHeight = 20;
                    colRange.Merge();
                    sheet.Cells[7 + i, colsname[j].Item1] = String.Format("数据行:{0}，列{1}", 7 + i, j + 1);

                    var border = colRange.Borders;
                    Utility.AddNativieResource(border);
                    border.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
            }

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
           "        开发人员按照团队成员的研发工作量占比排序\r\n" +
           "        测试人员按照团队成员的Bug产出率排序\r\n" +
           "        各明细工作量是实际填写的工作日志某类的工作量，和工作日志中分类一致\r\n" +
           "        Bug数：对于开发人员是本迭代被测试出的Bug总数\r\n" +
           "                 对于测试人员是本迭代测试出的Bug总数";

            var titleFont = titleRange.Font;
            Utility.AddNativieResource(titleFont);
            titleFont.Size = 11;

            var tmpchar = titleRange.Characters[1, 3];

            var tmpfont = tmpchar.Font;
            Utility.AddNativieResource(tmpchar);
            Utility.AddNativieResource(tmpfont);
            tmpfont.Bold = true;

            var tmpchar2 = titleRange.Characters[30, 45];

            var tmpfont2 = tmpchar2.Font;
            Utility.AddNativieResource(tmpchar2);
            Utility.AddNativieResource(tmpfont2);
            tmpfont2.Color = System.Drawing.Color.Red.ToArgb();

            ExcelInterop.Range titleRange2 = sheet.Range[sheet.Cells[10, "L"], sheet.Cells[10, "W"]];
            titleRange2.Merge();
            Utility.AddNativieResource(titleRange2);
            titleRange.Merge();
            sheet.Cells[10, "L"] = "      标准工作量：迭代天数*8\r\n" +
                                   "      实际投入工作量：不包含请假实际的工作量\r\n" +
                                   "      实际饱和度：实际投入工作量 / 标准工作量\r\n" +
                                   "      Bug产出率：bug数 / 实际投入工作量";
        }

        private int FillWorkloadData(List<WorkloadEntity> workloads, int startrow)
        {
            int firstrow = startrow;

            int standardWorkingDays = 0;
            IOrderedEnumerable<IGrouping<string, WorkloadEntity>> orderedLoads;
            int workloadCount = 0;
            GetOrderedWorkloads(workloads, out workloadCount, out standardWorkingDays, out orderedLoads);

            foreach (var load in orderedLoads)
            {
                string person = load.Key.Split(new char[] { '<' }, StringSplitOptions.RemoveEmptyEntries)[0];

                double leavetime = load.Where(wl => wl.SupperType == "请假").Sum(wl => wl.SumHours + wl.OverTimes);

                var dev = load.Where(wl => wl.SupperType == "研发");
                double devtime = dev.Sum(wl => wl.SumHours + wl.OverTimes);
                double devtime1 = dev.Where(wl => wl.Type == "开发").Sum(wl => wl.SumHours + wl.OverTimes);
                double devtime2 = dev.Where(wl => wl.Type == "需求").Sum(wl => wl.SumHours + wl.OverTimes);
                double devtime3 = dev.Where(wl => wl.Type == "设计").Sum(wl => wl.SumHours + wl.OverTimes);
                double devtime4 = dev.Where(wl => wl.Type == "测试设计").Sum(wl => wl.SumHours + wl.OverTimes);
                double devtime5 = dev.Where(wl => wl.Type == "测试执行").Sum(wl => wl.SumHours + wl.OverTimes);
                double devtime6 = devtime - (devtime1 + devtime2 + devtime3 + devtime4 + devtime5);

                double mgrtime = load.Where(wl => wl.SupperType == "管理").Sum(wl => wl.SumHours + wl.OverTimes);
                double oprtime = load.Where(wl => wl.SupperType == "运维").Sum(wl => wl.SumHours + wl.OverTimes);
                double doctime = load.Where(wl => wl.SupperType == "文档").Sum(wl => wl.SumHours + wl.OverTimes);
                double studytime = load.Where(wl => wl.SupperType == "学习交流").Sum(wl => wl.SumHours + wl.OverTimes);
                double presalestime = load.Where(wl => wl.SupperType == "售前/推广").Sum(wl => wl.SumHours + wl.OverTimes);
                double othertime = load.Sum(wl => wl.SumHours + wl.OverTimes) - (devtime + mgrtime + doctime + studytime + presalestime);

                sheet.Cells[startrow, "B"] = person;
                sheet.Cells[startrow, "C"] = standardWorkingDays * 8;
                sheet.Cells[startrow, "D"] = load.Sum(wl => wl.SumHours + wl.OverTimes) - leavetime;
                sheet.Cells[startrow, "E"] = leavetime;
                sheet.Cells[startrow, "F"] = String.Format("=D{0}/C{0}", startrow);
                sheet.Cells[startrow, "I"] = devtime / (load.Sum(wl => wl.SumHours + wl.OverTimes) - leavetime);
                sheet.Cells[startrow, "J"] = devtime1;
                sheet.Cells[startrow, "K"] = String.Format("=IF(J{0}<>0,J{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "L"] = devtime2;
                sheet.Cells[startrow, "M"] = String.Format("=IF(L{0}<>0,L{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "N"] = devtime3;
                sheet.Cells[startrow, "O"] = String.Format("=IF(N{0}<>0,N{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "P"] = devtime4;
                sheet.Cells[startrow, "Q"] = String.Format("=IF(P{0}<>0,P{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "R"] = devtime5;
                sheet.Cells[startrow, "S"] = String.Format("=IF(R{0}<>0,R{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "T"] = devtime6;
                sheet.Cells[startrow, "U"] = String.Format("=IF(T{0}<>0,T{0}/D{0}, \"\"", startrow);

                sheet.Cells[startrow, "V"] = mgrtime;
                sheet.Cells[startrow, "W"] = String.Format("=IF(V{0}<>0,V{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "X"] = oprtime;
                sheet.Cells[startrow, "Y"] = String.Format("=IF(X{0}<>0,X{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "Z"] = doctime;
                sheet.Cells[startrow, "AA"] = String.Format("=IF(Z{0}<>0,Z{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "AB"] = studytime;
                sheet.Cells[startrow, "AC"] = String.Format("=IF(AB{0}<>0,AB{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "AD"] = presalestime;
                sheet.Cells[startrow, "AE"] = String.Format("=IF(AD{0}<>0,AD{0}/D{0}, \"\"", startrow);
                sheet.Cells[startrow, "AF"] = othertime;
                sheet.Cells[startrow, "AG"] = String.Format("=IF(AF{0}<>0,AF{0}/D{0}, \"\"", startrow);

                startrow++;
            }

            ExcelInterop.Range devRange = sheet.Range[sheet.Cells[firstrow, "B"], sheet.Cells[startrow - 1, "AG"]];
            devRange.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignRight;
            devRange.WrapText = true;
            devRange.RowHeight = 20;
            devRange.ColumnWidth = 7;

            var borderDevRange = devRange.Borders;
            Utility.AddNativieResource(borderDevRange);
            borderDevRange.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;

            ExcelInterop.Range devRange2 = sheet.Range[sheet.Cells[firstrow, "B"], sheet.Cells[startrow - 1, "B"]];
            Utility.AddNativieResource(devRange2);
            devRange2.HorizontalAlignment = ExcelInterop.XlHAlign.xlHAlignCenter;

            ExcelInterop.Range devRange3 = sheet.Range[sheet.Cells[firstrow, "F"], sheet.Cells[startrow - 1, "F"]];
            Utility.AddNativieResource(devRange3);
            devRange3.NumberFormat = "#%";

            for (int col = 9; col <= 33; col += 2)
            {
                ExcelInterop.Range range = sheet.Range[sheet.Cells[firstrow, col], sheet.Cells[startrow - 1, col]];
                Utility.AddNativieResource(range);
                range.NumberFormat = "#%";
            }
            return startrow;
        }

        private void GetOrderedWorkloads(List<WorkloadEntity> workloads, out int workloadCount, out int standardWorkingDays, out IOrderedEnumerable<IGrouping<string, WorkloadEntity>> orderedLoads)
        {
            var loads = workloads.GroupBy(wl => wl.AssignedTo);
            var ite = TFS.Utility.GetBestIteration(project.Name);
            var daysoff = Iteration.GetProjectIterationDaysOff(this.project.Name, ite.Id);
            standardWorkingDays = (int)((DateTime.Parse(ite.EndDate) - DateTime.Parse(ite.StartDate)).TotalDays) + 1 - daysoff.Count;
            for (DateTime dt = DateTime.Parse(ite.StartDate); dt < DateTime.Parse(ite.EndDate).AddDays(1); dt = dt.AddDays(1))
            {
                if (dt.DayOfWeek == DayOfWeek.Sunday) standardWorkingDays--;//再刨掉礼拜天
            }

            orderedLoads = loads.OrderByDescending(
                load => (
                    load.Sum(wl => wl.SumHours + wl.OverTimes)
                    -
                    load.Where(wl => wl.SupperType == "请假").Sum(wl => wl.SumHours + wl.OverTimes)
                )
            );

            workloadCount = loads.Count();
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

            Utility.BuildFormalTableHeader(sheet, 11, "B", 13, "AG");

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
            startrow = FillWorkloadData(devload, startrow);
            startrow = FillWorkloadData(testload, startrow + 1);



            sheet.Cells[startrow, "B"] = "合计";
            sheet.Cells[startrow, "C"] = String.Format("=sum(C14:C{0}", startrow - 1);
            sheet.Cells[startrow, "D"] = String.Format("=sum(D14:D{0}", startrow - 1);
            sheet.Cells[startrow, "E"] = String.Format("=sum(E14:E{0}", startrow - 1);
            sheet.Cells[startrow, "F"] = String.Format("=IF(C{0}<>0,D{0}/C{0},\"\")", startrow);
            sheet.Cells[startrow, "G"] = String.Format("=sum(G14:G{0}", startrow - 1);

            sheet.Cells[startrow, "J"] = String.Format("=sum(J14:J{0}", startrow - 1);
            sheet.Cells[startrow, "K"] = String.Format("=J{0}/D{0}", startrow);
            sheet.Cells[startrow, "L"] = String.Format("=sum(L14:L{0}", startrow - 1);
            sheet.Cells[startrow, "M"] = String.Format("=L{0}/D{0}", startrow);
            sheet.Cells[startrow, "N"] = String.Format("=sum(N14:N{0}", startrow - 1);
            sheet.Cells[startrow, "O"] = String.Format("=N{0}/D{0}", startrow);
            sheet.Cells[startrow, "P"] = String.Format("=sum(P14:P{0}", startrow - 1);
            sheet.Cells[startrow, "Q"] = String.Format("=P{0}/D{0}", startrow);
            sheet.Cells[startrow, "R"] = String.Format("=sum(R14:R{0}", startrow - 1);
            sheet.Cells[startrow, "S"] = String.Format("=R{0}/D{0}", startrow);
            sheet.Cells[startrow, "T"] = String.Format("=sum(T14:T{0}", startrow - 1);
            sheet.Cells[startrow, "U"] = String.Format("=T{0}/D{0}", startrow);
            sheet.Cells[startrow, "V"] = String.Format("=sum(V14:V{0}", startrow - 1);
            sheet.Cells[startrow, "W"] = String.Format("=V{0}/D{0}", startrow);
            sheet.Cells[startrow, "X"] = String.Format("=sum(X14:X{0}", startrow - 1);
            sheet.Cells[startrow, "Y"] = String.Format("=X{0}/D{0}", startrow);
            sheet.Cells[startrow, "Z"] = String.Format("=sum(Z14:Z{0}", startrow - 1);
            sheet.Cells[startrow, "AA"] = String.Format("=Z{0}/D{0}", startrow);
            sheet.Cells[startrow, "AB"] = String.Format("=sum(AB14:AB{0}", startrow - 1);
            sheet.Cells[startrow, "AC"] = String.Format("=AB{0}/D{0}", startrow);
            sheet.Cells[startrow, "AD"] = String.Format("=sum(AD14:AD{0}", startrow - 1);
            sheet.Cells[startrow, "AE"] = String.Format("=AD{0}/D{0}", startrow);
            sheet.Cells[startrow, "AF"] = String.Format("=sum(AF14:AF{0}", startrow - 1);
            sheet.Cells[startrow, "AG"] = String.Format("=AF{0}/D{0}", startrow);

            sheet.Cells[7, "B"] = String.Format("=C{0}", startrow);
            sheet.Cells[7, "F"] = String.Format("=D{0}", startrow);
            sheet.Cells[7, "I"] = String.Format("=F7/B7", startrow);
            sheet.Cells[7, "K"] = String.Format("=J{0}+L{0}+N{0}+P{0}+R{0}+T{0}", startrow);
            sheet.Cells[7, "M"] = String.Format("=K7/F7", startrow);

            ExcelInterop.Range range2 = sheet.Range[sheet.Cells[startrow, "F"], sheet.Cells[startrow, "F"]];
            Utility.AddNativieResource(range2);
            range2.NumberFormat = "#%";

            for (int i = 9; i <= 33; i += 2)
            {
                SetCellPercentFormat(startrow, i);
            }

            SetCellPercentFormat(7, 9);
            SetCellPercentFormat(7, 13);

            #region 画图，暂时作废，藏起来
            //int chartstart = startrow + 2;

            //ExcelInterop.Range workloadChartRange = sheet.Range[sheet.Cells[chartstart, "B"], sheet.Cells[chartstart + 10, "AG"]];

            //ExcelInterop.ChartObjects charts = sheet.ChartObjects(Type.Missing) as ExcelInterop.ChartObjects;
            //Utility.AddNativieResource(charts);

            //ExcelInterop.ChartObject workloadChartObject = charts.Add(0, 0, workloadChartRange.Width, workloadChartRange.Height);
            //Utility.AddNativieResource(workloadChartObject);
            //ExcelInterop.Chart workloadChart = workloadChartObject.Chart;//设置图表数据区域。
            //Utility.AddNativieResource(workloadChart);

            ////=工作量统计!$B$14:$B$33,工作量统计!$F$14:$F$33
            //ExcelInterop.Range datasource = sheet.get_Range(String.Format("B14:B{0},F14:F{0}", startrow - 1));//不是："B14:B25","F14:F25"
            ////ExcelInterop.Range datasource = sheet.get_Range("B14:B25,F14:F25");//不是："B14:B25","F14:F25"
            //Utility.AddNativieResource(datasource);
            //workloadChart.ChartWizard(datasource, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, "工作量饱和度", "人员", "工作量", Type.Missing);
            //workloadChart.ApplyDataLabels();//图形上面显示具体的值
            ////将图表移到数据区域之下。
            //workloadChartObject.Left = Convert.ToDouble(workloadChartRange.Left);
            //workloadChartObject.Top = Convert.ToDouble(workloadChartRange.Top) + 20;

            //chartstart += 12;
            //ExcelInterop.Range bugChartRange = sheet.Range[sheet.Cells[chartstart, "B"], sheet.Cells[chartstart + 10, "K"]];
            //ExcelInterop.ChartObject bugChartObject = charts.Add(0, 0, bugChartRange.Width, bugChartRange.Height);
            //Utility.AddNativieResource(bugChartObject);
            //ExcelInterop.Chart bugChart = bugChartObject.Chart;//设置图表数据区域。
            //Utility.AddNativieResource(bugChart);

            ////=工作量统计!$B$14:$B$33,工作量统计!$F$14:$F$33
            //ExcelInterop.Range datasource2 = sheet.get_Range(String.Format("B{0}:B{1},G{0}:G{1}", 14 + devload.Count + 2 - 1, startrow - 1));//不是："B14:B25","F14:F25"
            ////ExcelInterop.Range datasource = sheet.get_Range("B14:B25,F14:F25");//不是："B14:B25","F14:F25"
            //Utility.AddNativieResource(datasource2);
            //bugChart.ChartWizard(datasource2, XlChartType.xlColumnClustered, Type.Missing, XlRowCol.xlColumns, 1, 1, false, "测试人员BUG数", "人员", "BUG数", Type.Missing);
            //bugChart.ApplyDataLabels();//图形上面显示具体的值
            ////将图表移到数据区域之下。
            //bugChartObject.Left = Convert.ToDouble(bugChartRange.Left);
            //bugChartObject.Top = Convert.ToDouble(bugChartRange.Top) + 20;
            #endregion 画图，暂时作废，藏起来
            return startrow + 2;
        }

        private void SetCellPercentFormat(int startrow, int i)
        {
            ExcelInterop.Range range = sheet.Range[sheet.Cells[startrow, i], sheet.Cells[startrow, i]];
            Utility.AddNativieResource(range);
            range.NumberFormat = "#%";
        }
    }
}
