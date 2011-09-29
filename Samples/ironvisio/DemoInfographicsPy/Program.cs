using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
            var chart7 = new IG.StripeGrid();
            chart7.Add("Feature1", "Theme1", "Scenario1");
            chart7.Add("Feature1", "Theme2", "Scenario1");
            chart7.Add("Feature1", "Theme3", "Scenario1");
            chart7.Add("Feature1", "Theme1", "Scenario2");
            chart7.Add("Feature1", "Theme2", "Scenario2");
            chart7.Add("Feature1", "Theme3", "Scenario2");
            chart7.Render(session.Page);
            session.ResizePageToFit(0.5, 0.5);
 
        }

        
    }
}
