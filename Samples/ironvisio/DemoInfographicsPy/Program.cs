using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using InfoGraphicsPy;
using IG = InfoGraphicsPy;
using VA=VisioAutomation;

namespace DemoInfographicsPy
{
    class Program
    {
        static void Main(string[] args)
        {

            var igs = new InfoGraphicsPy.Session();

            igs.NewDocument();

            TestDraw(igs);
        }

        public static void TestDraw(IG.Session session)
        {
            var CategoryLabels = new[] { "A", "B", "C", "D", "E" };
            var DataPoints = new IG.DataPoints(new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 });

            session.NewPage();
            var chart1 = new IG.PieSliceChart(DataPoints,CategoryLabels);
            chart1.Draw(session);
            session.ResizePageToFit(0.5,0.5);

            session.NewPage();
            var chart2 = new IG.VerticalBarChart(DataPoints, CategoryLabels);
            chart2.Draw(session);
            session.ResizePageToFit(0.5, 0.5);

            session.NewPage();
            var chart3 = new IG.DoughnutSliceChart(DataPoints, CategoryLabels);
            chart3.Draw(session);
            session.ResizePageToFit(0.5, 0.5);

            session.NewPage();
            var chart4 = new IG.HorizontalBarChart(DataPoints, CategoryLabels);
            chart4.Draw(session);
            session.ResizePageToFit(0.5, 0.5);

            session.NewPage();
            var chart5 = new IG.VerticalBarChart(DataPoints, CategoryLabels);
            chart5.Draw(session);
            session.ResizePageToFit(0.5, 0.5);

            session.NewPage();
            var chart6 = new IG.MonthGrid(2011,09);
            chart6.Render(session.Page);
            session.ResizePageToFit(0.5, 0.5);

            session.NewPage();
            var text = @"
Productivity, Batch,Multi-Format Export,
Productivity, Batch,Save/Reload Batch Settings,
Compelling Graphics, Colors,Improved Gradient Designer,
Compelling Graphics, Effects,2-Color Glow,
Compelling Graphics, Effects,Add Noise,
Compelling Graphics, Effects,Bleach,
Compelling Graphics, Effects,Burn,
Compelling Graphics, Effects,Tint,
Compelling Graphics, Effects,Blur, Motion|Gaussian
Compelling Graphics, Effects,Emboss,
Basics, Setup,Faster Install,
Basics, Setup,Updated Splash Screen,
Basics, Supportability,Logging During Batch,

";
            var chart7 = IG.CategoryChart.FromCSV("VizDraw Feature Map",text);
            chart7.ToUpper = true;

            chart7.Render(session.Page);
            session.ResizePageToFit(0.5, 0.5);
 
        }

    }
}
