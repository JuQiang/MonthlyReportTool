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


        public void BuildSlides()
        {            
            //var allbugs = (new TFS.Utility()).RetrieveAllBugsByDate("2017-10-01", "2017-11-01");
            //var teambugs = allbugs.GroupBy(wi => wi.SystemTeamProject);

            //foreach (var teambug in teambugs)
            //{
            //    if (teambug.Key == "Bugs") continue;
            //    Console.WriteLine(teambug.Key);
            //    var pptApplication = new PowerPointInterop.Application();
            //    // Create the Presentation File
            //    var pptPresentation = pptApplication.Presentations.Add();// (MsoTriState.msoTrue);

            //    // Create new Slide
            //    var slides = pptPresentation.Slides;
            //    var customLayout = pptPresentation.SlideMaster.CustomLayouts[PowerPointInterop.PpSlideLayout.ppLayoutText];

            //    int page = 1;
            //    //BuildSlide1(slides.AddSlide(page++, customLayout));
            //    //BuildSlide2(slides.AddSlide(page++, customLayout));

            //    //BuildAllBugsSlide(teambug, slides.AddSlide(page++, customLayout));
            //    //BuildProjectMaintananceBugsSlide(teambug, slides.AddSlide(page++, customLayout));
            //    BuildBugsByAreaPathSlide(teambug, slides.AddSlide(page++, customLayout));

            //    pptPresentation.SaveAs(@"c:\mrt\" + teambug.Key + ".pptx", PowerPointInterop.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            //    //pptPresentation.Close();
            //    //pptApplication.Quit();
            //}
         }

        private void PrepareBugs()
        {
            //云平台-二级部门-三级部门-...-N级部门-TeamProject
            //bugs库（维护库）上的bug，增加一个行政组织。然后行政组织与TP有关联关系。
            //这里就是按照关联关系组织数据，比如刘龙，下面挂4个项目。那么这里就把四个项目，按照刘龙的映射关系，归集起来。
        }

        private void BuildSlide1(PowerPointInterop.Slide slide)
        {
            var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
            // Add title
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = "说  明";
            title.Font.NameFarEast = "微软雅黑";
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = 0x00C07000;
            title.Font.Size = 32;


            //日了狗了，所有的资料都是\n，其实特么的是\r\n，否则第二个indent level=2没效果！
            List<MyParagraph> list = new List<MyParagraph>();
            list.Add(new MyParagraph() { Text = "产品质量分析会议时间", ChildParagraphList = new List<MyParagraph>() { new MyParagraph() { Text = "与研发月度运营会议一起进行" } } });
            list.Add(new MyParagraph() { Text = "材料准备", ChildParagraphList = new List<MyParagraph>() { new MyParagraph() { Text = "产品质量分析报告（参考产品质量分析报告模板）：各研发部负责" } } });
            list.Add(new MyParagraph()
            {
                Text = "要求",
                ChildParagraphList = new List<MyParagraph>() {
                    new MyParagraph() { Text = "1. 除特殊说明， Bug统计均为项目库和维护库所有的bug" },
                    new MyParagraph() { Text = "2. 需要将字体为斜体部分的内容替换成实际的内容" },
                    new MyParagraph() { Text = "3. 报告中的图、表需根据实际情况进行统计替换" }
                }
            });

            var paralist = slide.Shapes[2].TextFrame.TextRange;
            foreach (var mypara in list)
            {
                var para = paralist.Paragraphs().InsertAfter(mypara.Text + Environment.NewLine);
                para.Font.NameFarEast = "微软雅黑";
                para.Font.Color.RGB = 0x00C07000;
                para.Font.Size = 18;

                foreach (var mysubpara in mypara.ChildParagraphList)
                {
                    var parasub = paralist.Paragraphs().InsertAfter(mysubpara.Text + Environment.NewLine);
                    parasub.Font.NameFarEast = "微软雅黑";
                    parasub.IndentLevel = 2;
                    parasub.Font.Color.RGB = 0x007F7F7F;
                    parasub.Font.Size = 16;
                }
            }

            //var parasub2 = objText2.Paragraphs().InsertAfter("产品质量分析报告（参考产品质量分析报告模板）：各研发部负责\r\n");
            //parasub2.IndentLevel = 2; parasub2.Font.Color.RGB = 0x007F7F7F;
            //parasub2.Characters(9, 14).Font.Color.RGB = 0x000000FF;
            //parasub2.Font.Size = 16;

            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "This demo is created by FPPT using C# - Download free templates from http://FPPT.com";
        }

        private void BuildSlide2(PowerPointInterop.Slide slide)
        {
            var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
            // Add title
            var title = slide.Shapes[1].TextFrame.TextRange;
            title.Text = "目  录";
            title.Font.NameFarEast = "微软雅黑";
            title.Font.Bold = MsoTriState.msoTrue;
            title.Font.Color.RGB = 0x00C07000;
            title.Font.Size = 32;

            var shape2 = slide.Shapes[2];
            float width = slide.Shapes[1].Width / 2;
            shape2.Width = width;

            var shape3 = shape2.Duplicate();
            shape3.Top = shape2.Top;
            shape3.Left = shape2.Left + width;

            //日了狗了，所有的资料都是\n，其实特么的是\r\n，否则第二个indent level=2没效果！
            List<MyParagraph> list = new List<MyParagraph>();
            list.Add(new MyParagraph()
            {
                Text = "一、Bug数量及分布情况统计分析",
                ChildParagraphList = new List<MyParagraph>() {
                    new MyParagraph() { Text = "1. Bug数量及类别分布情况分析" },
                    new MyParagraph() { Text = "2. Bug项目库、维护库分布情况分析" },
                    new MyParagraph() { Text = "3. Bug区域分布情况分析" },
                    new MyParagraph() { Text = "4. Bug按开发人员统计分析" },
                    new MyParagraph() { Text = "5. Bug问题等级分布情况分析" },
                    new MyParagraph() { Text = "6. 预警问题分析" },
                }
            });
            list.Add(new MyParagraph()
            {
                Text = "二、Bug修复情况",
                ChildParagraphList = new List<MyParagraph>() {
                    new MyParagraph() { Text = "1. Bug状态统计、未关闭bug的原因分析" },
                    new MyParagraph() { Text = "2. Bug项目库、维护库分布情况分析" },
                }
            });

            var paralist = shape2.TextFrame.TextRange;
            foreach (var mypara in list)
            {
                var para = paralist.Paragraphs().InsertAfter(mypara.Text + Environment.NewLine);
                para.Font.NameFarEast = "微软雅黑";
                para.Font.Color.RGB = 0x00C07000;
                para.Font.Bold = MsoTriState.msoTrue;
                para.Font.Size = 20;

                foreach (var mysubpara in mypara.ChildParagraphList)
                {
                    var parasub = paralist.Paragraphs().InsertAfter(mysubpara.Text + Environment.NewLine);
                    parasub.Font.NameFarEast = "微软雅黑";
                    parasub.IndentLevel = 2;
                    parasub.Font.Color.RGB = 0x00C07000;
                    parasub.Font.Bold = MsoTriState.msoTrue;
                    parasub.Font.Size = 18;
                }
            }

            List<MyParagraph> list2 = new List<MyParagraph>();
            list2.Add(new MyParagraph()
            {
                Text = "三、Bug原因分析",
                ChildParagraphList = new List<MyParagraph>() {
                    new MyParagraph() { Text = "1. 1、2级Bug原因分析" },
                    new MyParagraph() { Text = "2. 程序错误问题分析" },
                }
            });
            list2.Add(new MyParagraph()
            {
                Text = "四、本月工作量分布情况",
            });
            list2.Add(new MyParagraph()
            {
                Text = "五、已移除和已中止提交单分析",
            });
            list2.Add(new MyParagraph()
            {
                Text = "六、改进措施",
                ChildParagraphList = new List<MyParagraph>() {
                    new MyParagraph() { Text = "1. 上月改进措施执行情况" },
                    new MyParagraph() { Text = "2. 本月改进措施" },
                }
            });

            var paralist2 = shape3.TextFrame.TextRange;
            foreach (var mypara in list2)
            {
                var para = paralist2.Paragraphs().InsertAfter(mypara.Text + Environment.NewLine);
                para.Font.NameFarEast = "微软雅黑";
                para.Font.Color.RGB = 0x00C07000;
                para.Font.Bold = MsoTriState.msoTrue;
                para.Font.Size = 20;

                if (null == mypara.ChildParagraphList) continue;
                foreach (var mysubpara in mypara.ChildParagraphList)
                {
                    var parasub = paralist2.Paragraphs().InsertAfter(mysubpara.Text + Environment.NewLine);
                    parasub.Font.NameFarEast = "微软雅黑";
                    parasub.IndentLevel = 2;
                    parasub.Font.Color.RGB = 0x00C07000;
                    parasub.Font.Bold = MsoTriState.msoTrue;
                    parasub.Font.Size = 18;
                }
            }
            //var parasub2 = objText2.Paragraphs().InsertAfter("产品质量分析报告（参考产品质量分析报告模板）：各研发部负责\r\n");
            //parasub2.IndentLevel = 2; parasub2.Font.Color.RGB = 0x007F7F7F;
            //parasub2.Characters(9, 14).Font.Color.RGB = 0x000000FF;
            //parasub2.Font.Size = 16;

            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "This demo is created by FPPT using C# - Download free templates from http://FPPT.com";
        }

        //private void BuildAllBugsSlide(IGrouping<string,BugEntity> teambugs, PowerPointInterop.Slide slide)
        //{
        //    var frame = (slide.Shapes[1] as PowerPointInterop.Shape).TextFrame;
        //    // Add title
        //    var title = slide.Shapes[1].TextFrame.TextRange;
        //    title.Text = "一、Bug数量及分布情况统计分析 - 总体";
        //    title.Font.NameFarEast = "微软雅黑";
        //    title.Font.Bold = MsoTriState.msoTrue;
        //    title.Font.Color.RGB = 0x00C07000;
        //    title.Font.Size = 24;

        //    //all bugs: 958a895b-5c5f-4ce5-ac60-84f4b652941c
        //    //http://tfs.teld.cn:8080/tfs/Teld/OrgPortal/_workItems/resultsById/8cf850a5-8352-4492-9f52-e02a7c5ef6b7
            
        //    //{[wiql, 
        //    //select [System.Id], [System.CreatedDate], [System.CreatedBy], [System.Title], [System.AssignedTo], [System.State], 
        //    //[Teld.Bug.ResolvedReason], [Teld.Bug.Type], [Microsoft.VSTS.Common.Severity], [Teld.Bug.Verificator], [System.AreaPath], 
        //    //[Teld.Bug.PlanReleaseTime], [Teld.Bug.HopeFixSubmitTime], [Teld.Bug.IfAgreeResultState], [Teld.Bug.SubmitToAuditDate], 
        //    //[Teld.Bug.RepeatBugId], [System.TeamProject] 
        //    //from WorkItems 
        //    //where [System.WorkItemType] = 'Bug' and[System.TeamProject] <> 'Bugs' and[System.CreatedDate] <= '2017-12-31T00:00:00.0000000' 
        //    //and[System.CreatedDate] >= '2017-07-01T00:00:00.0000000' and[System.State] <> '已移除' order by[Teld.Bug.Type]]}
        //    var bugsdate = teambugs.GroupBy(wi => wi.CreatedYearMonth);
        //    List<string> dateseries = new List<string>();
        //    foreach (var date in bugsdate)
        //    {
        //        dateseries.Add(date.Key);
        //    }


        //    var bugstype = teambugs.GroupBy(wi => wi.TeldBugType);

        //    var bugtbl = slide.Shapes.AddTable(dateseries.Count + 2, bugstype.Count() * 2 + 2);

        //    bugtbl.Table.Cell(1, 1).Merge(bugtbl.Table.Cell(2, 1));
        //    var month = bugtbl.Table.Cell(1, 1).Shape.TextFrame.TextRange;
        //    month.Text = "月份";month.Font.Size = 12;

        //    int row = 3;
        //    foreach (string yearmonth in GetDateSeriesByFriendlyFormat(dateseries))
        //    {
        //        var dt = bugtbl.Table.Cell(row++, 1).Shape.TextFrame.TextRange;
        //        dt.Text = yearmonth;
        //        dt.Font.Size = 12;
        //    }

        //    int col = 2;
        //    foreach (var bugtype in bugstype)
        //    {
        //        bugtbl.Table.Cell(1, col).Merge(bugtbl.Table.Cell(1, col+1));
        //        var tr1 = bugtbl.Table.Cell(1, col).Shape.TextFrame.TextRange;                
        //        tr1.Text = bugtype.Key;tr1.Font.Size = 12;

        //        var tr2 = bugtbl.Table.Cell(2, col).Shape.TextFrame.TextRange;
        //        tr2.Text = "数量"; tr2.Font.Size = 12;
        //        var tr3 = bugtbl.Table.Cell(2, col + 1).Shape.TextFrame.TextRange;
        //        tr3.Text = "占比"; tr3.Font.Size = 12;

        //        row = 3;
        //        foreach (string yearmonth in dateseries)
        //        {
        //            var monthbugtype = bugtype.Where(wi => wi.CreatedYearMonth == yearmonth);
        //            var data = bugtbl.Table.Cell(row++, col).Shape.TextFrame.TextRange;
        //            data.Text = monthbugtype.Count().ToString();                    
        //            data.Font.Size = 12;
        //            data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;
        //        }

        //        col += 2;
        //    }
        //    bugtbl.Table.Cell(1, col).Merge(bugtbl.Table.Cell(2, col));
        //    var sumr= bugtbl.Table.Cell(1, col).Shape.TextFrame.TextRange;
        //    sumr.Text = "合计"; sumr.Font.Size = 12;

        //    for (row = 3; row < dateseries.Count + 3; row++)
        //    {
        //        int sum = 0;
        //        for (col = 2; col < 2 * bugstype.Count() + 2; col += 2)
        //        {
        //            sum += Convert.ToInt32(bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange.Text);
        //        }
        //        var data = bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange;
        //        data.Text = Convert.ToString(sum);
        //        data.Font.Size = 12;
        //        data.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;

                
        //        for (col = 2; col < 2 * bugstype.Count() + 2; col += 2)
        //        {
        //            var percent = bugtbl.Table.Cell(row, col + 1).Shape.TextFrame.TextRange;
        //            var num = Convert.ToInt32(bugtbl.Table.Cell(row, col).Shape.TextFrame.TextRange.Text);
        //            percent.Text = Convert.ToString(num * 100 / sum) + "%";
        //            percent.Font.Size = 12;
        //            percent.ParagraphFormat.Alignment = PowerPointInterop.PpParagraphAlignment.ppAlignRight;
        //        }
        //    }

        //    //bugtbl.Table.ScaleProportionally(0.5f);
        //}

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

        private string[] GetDateSeriesByFriendlyFormat(List<string> dateseries)
        {
            string[] array = new string[dateseries.Count];
            //显示去冗余用
            dateseries.CopyTo(array);
            bool isSameYear = true;
            string year = array[0].Substring(0, 4);
            for (int i = 1; i < array.Length; i++)
            {
                if (year != array[1].Substring(0, 4))
                {
                    isSameYear = false;
                    break;
                }
            }

            if (isSameYear)
            {
                for (int i = 0; i < array.Length; i++)
                {
                    array[i] = array[i].Substring(5);
                }
            }

            return array;
        }
        internal class MyParagraph
        {
            public string Text;
            public List<MyParagraph> ChildParagraphList;
        }
    }
}
