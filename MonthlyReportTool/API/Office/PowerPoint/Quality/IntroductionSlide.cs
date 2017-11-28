using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MonthlyReportTool.API.TFS.TeamProject;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace MonthlyReportTool.API.Office.PowerPoint.Quality
{
    public class IntroductionSlide :  PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public IntroductionSlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project, string yearmonth)
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
    }
}
