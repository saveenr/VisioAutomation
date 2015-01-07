using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingDrawTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Draw_OrgChart()
        {
            var client = GetScriptingClient();
            draw_org_chart(client, TestVisioAutomation.Properties.Resources.sampleorgchart1);
            client.Document.Close(true);
            VA.Documents.DocumentHelper.ForceCloseAll(client.VisioApplication.Documents);
        }

        private void draw_org_chart(VA.Scripting.Client client, string text)
        {
            var xmldoc = SXL.XDocument.Parse(text);
            var orgchart = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(client, xmldoc);
            client.Draw.OrgChart(orgchart);
        }

        [TestMethod]
        public void Scripting_Draw_DataTable()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var items = new[]
                {
                    new {Name = "X", Age = 28, Score = 16},
                    new {Name = "Y", Age = 32, Score = 23},
                    new {Name = "Z", Age = 45, Score = 12},
                    new {Name = "U", Age = 48, Score = 10}
                };

            var dt = new System.Data.DataTable();
            dt.Columns.Add("X", typeof (string));
            dt.Columns.Add("Age", typeof (int));
            dt.Columns.Add("Score", typeof (int));

            foreach (var item in items)
            {
                dt.Rows.Add(item.Name, item.Age, item.Score);
            }

            var widths = new[] {2.0, 1.5, 1.0};
            var heights = Enumerable.Repeat(0.25, items.Length).ToList();
            var shapes = client.Draw.Table(dt, widths, heights, new VA.Drawing.Size(0, 0));

            Assert.AreEqual(items.Length*3, shapes.Count);

            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_Grid()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var cellsize = new VA.Drawing.Size(0.5, 0.25);
            int cols = 3;
            int rows = 6;

            client.Document.OpenStencil("basic_u.vss");
            string stencil = "basic_u.vss";
            string mastername = "Rectangle";

            var stencildoc = client.Document.Get(stencil);
            var master = client.Master.Get(mastername, stencildoc);

            var grid = new VA.Models.Grid.GridLayout(cols, rows, cellsize, master);
            grid.Origin = new VA.Drawing.Point(0, 4);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_RectangleLineOval_0()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var shape_rect = client.Draw.Rectangle(1, 1, 3, 3);
            var shape_line = client.Draw.Line(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.Oval(0.2, 1, 3.8, 2);
            var shape_oval2 = client.Draw.Oval(new VA.Drawing.Point(2, 2), 0.5);

            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_BezierPolyLine_0()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var points = new[]
                {
                    new VA.Drawing.Point(0, 0),
                    new VA.Drawing.Point(2, 0.5),
                    new VA.Drawing.Point(2, 2),
                    new VA.Drawing.Point(3, 0.5)
                };

            var shape_bezier = client.Draw.Bezier(points);
            var shape_polyline = client.Draw.PolyLine(points);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_PieSlice()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var center = new VA.Drawing.Point(2, 2);
            double radius = 1.0;
            double start_angle = 0;
            double end_angle = System.Math.PI;

            var shape = client.Draw.PieSlice(center, radius, start_angle, end_angle);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_PieChart()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var center = new VA.Drawing.Point(2, 2);
            double radius = 1.0;
            var chart = new VA.Models.Charting.PieChart(center,radius);
            chart.DataPoints.Add(new VA.Models.Charting.DataPoint(1.0));
            chart.DataPoints.Add(new VA.Models.Charting.DataPoint(2.0));
            chart.DataPoints.Add(new VA.Models.Charting.DataPoint(3.0));
            chart.DataPoints.Add(new VA.Models.Charting.DataPoint(4.0));
            client.Draw.PieChart(chart);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_BarChart()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var rect1 = new VA.Drawing.Rectangle(0, 0, 4, 4);
            var chart1 = new VA.Models.Charting.BarChart(rect1);
            chart1.DataPoints.Add(new VA.Models.Charting.DataPoint(1.0));
            chart1.DataPoints.Add(new VA.Models.Charting.DataPoint(2.0));
            chart1.DataPoints.Add(new VA.Models.Charting.DataPoint(3.0));
            chart1.DataPoints.Add(new VA.Models.Charting.DataPoint(4.0));
            client.Draw.BarChart(chart1);

            var rect2 = new VA.Drawing.Rectangle(5, 0, 9, 4);
            var chart2= new VA.Models.Charting.BarChart(rect2);
            chart2.DataPoints.Add(new VA.Models.Charting.DataPoint(1.0));
            chart2.DataPoints.Add(new VA.Models.Charting.DataPoint(2.0));
            chart2.DataPoints.Add(new VA.Models.Charting.DataPoint(-3.0));
            chart2.DataPoints.Add(new VA.Models.Charting.DataPoint(4.0));
            client.Draw.BarChart(chart2);

            var rect3 = new VA.Drawing.Rectangle(10, 0, 14, 4);
            var chart3 = new VA.Models.Charting.BarChart(rect3);
            chart3.DataPoints.Add(new VA.Models.Charting.DataPoint(-1.0));
            chart3.DataPoints.Add(new VA.Models.Charting.DataPoint(-2.0));
            chart3.DataPoints.Add(new VA.Models.Charting.DataPoint(-3.0));
            chart3.DataPoints.Add(new VA.Models.Charting.DataPoint(-4.0));
            client.Draw.BarChart(chart3);

            client.Page.Get().ResizeToFitContents(new VA.Drawing.Size(1.0,1.0));
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_AreaChart()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var rect1 = new VA.Drawing.Rectangle(0, 0, 4, 4);
            var chart1 = new VA.Models.Charting.AreaChart(rect1);
            chart1.DataPoints.Add(1.0);
            chart1.DataPoints.Add(2.0);
            chart1.DataPoints.Add(3.0);
            chart1.DataPoints.Add(4.0);
            client.Draw.AreaChart(chart1);

            var rect2 = new VA.Drawing.Rectangle(5, 0, 9, 4);
            var chart2 = new VA.Models.Charting.AreaChart(rect2);
            chart2.DataPoints.Add(1.0);
            chart2.DataPoints.Add(2.0);
            chart2.DataPoints.Add(-3.0);
            chart2.DataPoints.Add(4.0);
            client.Draw.AreaChart(chart2);

            var rect3 = new VA.Drawing.Rectangle(10, 0, 14, 4);
            var chart3 = new VA.Models.Charting.AreaChart(rect3);
            chart3.DataPoints.Add(-1.0);
            chart3.DataPoints.Add(-2.0);
            chart3.DataPoints.Add(-3.0);
            chart3.DataPoints.Add(-4.0);
            client.Draw.AreaChart(chart3);

            client.Page.Get().ResizeToFitContents(new VA.Drawing.Size(1.0, 1.0));
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_DirectedGraph1()
        {
            var client = GetScriptingClient();
            draw_dg(client, TestVisioAutomation.Properties.Resources.sampleflowchart1);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_DirectedGraph2()
        {
            var client = GetScriptingClient();
            draw_dg(client, TestVisioAutomation.Properties.Resources.sampleflowchart2);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_DirectedGraph3()
        {
            var client = GetScriptingClient();
            draw_dg(client, TestVisioAutomation.Properties.Resources.sampleflowchart3);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_DirectedGraph4()
        {
            var client = GetScriptingClient();
            draw_dg(client, TestVisioAutomation.Properties.Resources.sampleflowchart4);
            client.Document.Close(true);
        }

        private void draw_dg(VA.Scripting.Client client, string dg_text)
        {
            var dg_xml = SXL.XDocument.Parse(dg_text);
            var dg_model = VA.Scripting.DirectedGraph.DirectedGraphBuilder.LoadFromXML(client, dg_xml);

            // this is a temporary fix to handle the fact that server_u.vss in Visio 2013 doesn't result in server_u.vssx 
            // gettign automatically loaded

            var version = client.Application.Version;
            if (version.Major >= 15)
            {
                foreach (var drawing in dg_model)
                {
                    foreach (var shape in drawing.Shapes)
                    {
                        if (shape.StencilName == "server_u.vss")
                        {
                            shape.StencilName = "server_u.vssx";
                        }
                    }
                }
            }
            
            client.Draw.DirectedGraph(dg_model);
        }
        
        [TestMethod]
        public void Scripting_Drop_Master()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            string basic_stencil_name = "Basic_U.VSS";
            var basic_stencil = client.Document.OpenStencil(basic_stencil_name);
            var master = client.Master.Get("Rectangle", basic_stencil);
            client.Master.Drop(master, 2, 2);
            var application = client.VisioApplication;
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            Assert.AreEqual(1, shapes.Count);
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Drop_Many()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(10, 10), false);
            var basic_stencil = client.Document.OpenStencil("Basic_U.VSS");

            var m1 = client.Master.Get("Rectangle", basic_stencil);
            var m2 = client.Master.Get("Ellipse", basic_stencil);

            var masters = new[] {m1, m2};
            var points = VA.Drawing.Point.FromDoubles(new[] { 1.0, 2.0, 3.0, 4.0, 1.5, 4.5, 5.7, 2.4 }).ToList();

            client.Master.Drop(masters, points);
            
            Assert.AreEqual(4, client.VisioApplication.ActivePage.Shapes.Count);
            client.Document.Close(true);
        }
    }
}