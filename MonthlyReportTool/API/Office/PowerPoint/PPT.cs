using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
//using ExcelInterop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeInterop = Microsoft.Office.Core;
using Microsoft.Office.Core;
using MonthlyReportTool.API.TFS.WorkItem;

namespace MonthlyReportTool.API.Office.PowerPoint
{
    public class PPT
    {
        //        public void Test()
        //        {
        //            String strTemplate, strPic;
        //            strTemplate =@"c:\mrt\产品质量分析报告模板（20171103更新）.potx";
        //            strPic = @"c:\mrt\juqiang.jpg";

        //            PowerPointInterop.Application objApp;
        //            PowerPointInterop.Presentations objPresSet;
        //            PowerPointInterop._Presentation objPres;
        //            PowerPointInterop.Slides objSlides;
        //            PowerPointInterop._Slide objSlide;
        //            PowerPointInterop.TextRange objTextRng;
        //            PowerPointInterop.Shapes objShapes;
        //            PowerPointInterop.Shape objShape;
        //            PowerPointInterop.SlideShowWindows objSSWs;
        //            PowerPointInterop.SlideShowTransition objSST;
        //            PowerPointInterop.SlideShowSettings objSSS;
        //            PowerPointInterop.SlideRange objSldRng;
        ////            GraphInterop.Chart objChart;

        //            //Create a new presentation based on a template.
        //            objApp = new PowerPointInterop.Application();
        //            objApp.Visible = MsoTriState.msoTrue;
        //            objPresSet = objApp.Presentations;
        //            objPres = objPresSet.Open(strTemplate,MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
        //            objSlides = objPres.Slides;


        //            //Build Slide #1:
        //            //Add text to the slide, change the font and insert/position a 
        //            //picture on the first slide.
        //            objSlide = objSlides.Add(1, PowerPointInterop.PpSlideLayout.ppLayoutTitleOnly);
        //            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
        //            objTextRng.Text = "My Sample Presentation";
        //            objTextRng.Font.Name = "Comic Sans MS";
        //            objTextRng.Font.Size = 48;
        //            objSlide.Shapes.AddPicture(strPic, MsoTriState.msoFalse, MsoTriState.msoTrue,
        //            150, 150, 500, 350);

        //            //Build Slide #2:
        //            //Add text to the slide title, format the text. Also add a chart to the
        //            //slide and change the chart type to a 3D pie chart.
        //            objSlide = objSlides.Add(2, PowerPointInterop.PpSlideLayout.ppLayoutTitleOnly);
        //            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
        //            objTextRng.Text = "My Chart";
        //            objTextRng.Font.Name = "Comic Sans MS";
        //            objTextRng.Font.Size = 48;
        //            //objChart = (GraphInterop.Chart)objSlide.Shapes.AddOLEObject(150, 150, 480, 320,
        //            //"MSGraph.Chart.8", "", MsoTriState.msoFalse, "", 0, "",
        //            //MsoTriState.msoFalse).OLEFormat.Object;
        //            //objChart.ChartType = GraphInterop.XlChartType.xl3DPie;
        //            //objChart.Legend.Position = GraphInterop.XlLegendPosition.xlLegendPositionBottom;
        //            //objChart.HasTitle = true;
        //            //objChart.ChartTitle.Text = "Here it is...";

        //            //Build Slide #3:
        //            //Change the background color of this slide only. Add a text effect to the slide
        //            //and apply various color schemes and shadows to the text effect.
        //            objSlide = objSlides.Add(3, PowerPointInterop.PpSlideLayout.ppLayoutBlank);
        //            objSlide.FollowMasterBackground = MsoTriState.msoFalse;
        //            objShapes = objSlide.Shapes;
        //            objShape = objShapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect27,
        //              "The End", "Impact", 96, MsoTriState.msoFalse, MsoTriState.msoFalse, 230, 200);

        //            //var slide0 = objSlides.FindBySlideID(1);
        //            //Modify the slide show transition settings for all 3 slides in
        //            //the presentation.
        //            int[] SlideIdx = new int[3];
        //            for (int i = 0; i < 3; i++) SlideIdx[i] = i + 1;
        //            objSldRng = objSlides.Range(SlideIdx);
        //            objSST = objSldRng.SlideShowTransition;
        //            objSST.AdvanceOnTime = MsoTriState.msoTrue;
        //            objSST.AdvanceTime = 3;
        //            objSST.EntryEffect = PowerPointInterop.PpEntryEffect.ppEffectBoxOut;

        //            ////Prevent Office Assistant from displaying alert messages:
        //            //bAssistantOn = objApp.Assistant.On;
        //            //objApp.Assistant.On = false;

        //            //Run the Slide show from slides 1 thru 3.
        //            objSSS = objPres.SlideShowSettings;
        //            objSSS.StartingSlide = 1;
        //            objSSS.EndingSlide = 3;
        //            objSSS.Run();

        //            //Wait for the slide show to end.
        //            objSSWs = objApp.SlideShowWindows;
        //            while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(100);

        //            //Reenable Office Assisant, if it was on:
        //            //if (bAssistantOn)
        //            //{
        //            //    objApp.Assistant.On = true;
        //            //    objApp.Assistant.Visible = false;
        //            //}

        //            //Close the presentation without saving changes and quit PowerPoint.
        //            objPres.Close();
        //            objApp.Quit();
        //        }


        private void PrepareBugs()
        {
            //云平台-二级部门-三级部门-...-N级部门-TeamProject
            //bugs库（维护库）上的bug，增加一个行政组织。然后行政组织与TP有关联关系。
            //这里就是按照关联关系组织数据，比如刘龙，下面挂4个项目。那么这里就把四个项目，按照刘龙的映射关系，归集起来。
        }


        private void BuildProjectMaintananceBugsSlide(IGrouping<string, BugEntity> teambugs, PowerPointInterop.Slide slide)
        {
            var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
            // Add title
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = "一、Bug数量及分布情况统计分析 - 项目库与维护库";
            title.Font.NameFarEast = "微软雅黑";
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = 0x00C07000;
            title.Font.Size = 24;

        }

        //private void BuildBugsByAreaPathSlide(IGrouping<string, BugEntity> teambugs, PowerPointInterop.Slide slide)
        //{
        //    var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
        //    // Add title
        //    var title = slide.Shapes[1].TextFrame.TextRange;
        //    title.Text = "一、Bug数量及分布情况统计分析 - 区域路径";
        //    title.Font.NameFarEast = "微软雅黑";
        //    title.Font.Bold = MsoTriState.msoTrue;
        //    title.Font.Color.RGB = 0x00C07000;
        //    title.Font.Size = 24;

        //    var pathbugs = teambugs.GroupBy(wi => wi.SystemAreaPath);
        //    #region 找到二级分类
        //    //FCP，一级分类
        //    //FCP   5，这个是没登记到具体二级分类的
        //    //FCP\02业务运营平台\04销售管理   8，这个应该和所有02业务运营平台的合并
        //    //FCP\01基础服务平台\01业务公共   7，这个应该和所有01业务公共的合并，一下雷同
        //    //FCP\01基础服务平台   1
        //    //FCP\02业务运营平台\01代金券管理   3
        //    //FCP\02业务运营平台\06收退款管理   2
        //    //FCP\01基础服务平台\02业务数据\03SAP启用设置   2
        //    //FCP\01基础服务平台\04企业信息管理   8
        //    //FCP\01基础服务平台\02业务数据   5
        //    //FCP\01基础服务平台\03客户管理   6
        //    List<string> subgroup = new List<string>();//二级分类
        //    string mainkey = teambugs.Key;
        //    foreach (var path in pathbugs)
        //    {
        //        Console.WriteLine(path.Key + "     " + path.Count());
        //        if (path.Key == mainkey) { subgroup.Add(path.Key); continue; }//未登记到明细二级问题的。
        //        string tmp = path.Key.Replace(mainkey + "\\", "");

        //        int pos = tmp.IndexOf("\\");
        //        if (pos < 0) {
        //            if (subgroup.Contains(tmp)) continue;
        //            subgroup.Add(tmp); continue; }
        //        else
        //        {
        //            if (subgroup.Contains(tmp.Substring(0, pos))) continue;
        //            subgroup.Add(tmp.Substring(0, pos));
        //            continue;
        //        }
        //    }
        //    subgroup = subgroup.OrderBy(wi => wi).ToList();
        //    #endregion 找到二级分类

        //    #region 按照二级分类汇总
        //    Dictionary<string, int> results = new Dictionary<string, int>();
        //    foreach (string sub in subgroup)
        //    {
                
        //        results.Add(sub, 0);
        //    }
        //    foreach (var pathbug in pathbugs){
        //        foreach (string sub in subgroup)
        //        {
        //            if (pathbug.Key==mainkey)
        //            {
        //                results[mainkey] = pathbug.Count();
        //            }
        //            else if (pathbug.Key.StartsWith(mainkey + "\\" + sub))
        //            {
        //                results[sub] = results[sub] + pathbug.Count();
        //            }
        //        }
        //    }

        //    Console.WriteLine();
        //    foreach (string key in results.Keys) { Console.WriteLine(key + "    " + results[key]); }
        //    Console.WriteLine("=================================");
        //    #endregion 按照二级分类汇总

        //    #region 画柱状图

        //    var chart = slide.Shapes.AddChart2(-1, OfficeInterop.XlChartType.xlBarClustered, 100, 100, 100, 100).Chart;
        //    //input data
        //    //chart.ChartData.Activate();
        //    //chart.ChartData[0, 0].Text = "";
        //    //int num = kinds.Count();
        //    //Excel.Workbook workbook = chart.ChartData.Workbook;
        //    //Excel.Worksheet sheet = chart.ChartData.Workbook.Worksheets["Sheet1"];
        //    //sheet.Cells.Clear();

        //    //Excel.Range range;
        //    //object[] objHeaders = { "数量", "数据1" };
        //    //range = sheet.get_Range("A1", "B1");
        //    //range.set_Value(Type.Missing, objHeaders);

        //    //var data = new object[num, 2];
        //    //foreach (int n in Enumerable.Range(0, num))
        //    //{
        //    //    data[n, 0] = kinds[n];
        //    //    data[n, 1] = values[n];
        //    //}

        //    //range = sheet.get_Range("A2", "B" + (num + 1));
        //    //sheet.get_Range("A2", "B" + (num + 1)).Value = data;
        //    //sheet.get_Range("B1").Value = title;
        //    ////Excel.Range chartRange = sheet.get_Range("A1", "B5");
        //    //chart.ChartWizard(sheet.get_Range("A1", "B5"), missing, missing, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns, 1, 1, true, "多糖商品销量分析", "月份", "销量", missing);

        //    ////set different part's color
        //    //for (int i = 1; i < 2; i++)
        //    //{
        //    //    PPT.Series series = chart.SeriesCollection(i);
        //    //    for (int j = 1; j <= num; j++)
        //    //    {
        //    //        PPT.Point point = series.Points(j);
        //    //        point.Format.Fill.ForeColor.RGB = color[j - 1];
        //    //    }
        //    //}
        //    #endregion 画柱状图
        //}

        
        
    }
}
