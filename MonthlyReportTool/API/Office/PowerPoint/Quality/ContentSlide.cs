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
    public class ContentSlide :  PowerPointSlideBase, IPowerPointQualitySlide
    {
        private PowerPointInterop.Slide slide;
        public ContentSlide(PowerPointInterop.Slide slide) : base(slide)
        {
            this.slide = slide;
        }
        public void Build(ProjectEntity project, string yearmonth)
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
    }
}
