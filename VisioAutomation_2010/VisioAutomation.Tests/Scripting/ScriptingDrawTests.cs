using System;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Models.Layouts.Grid;
using VisioAutomation.Scripting.Models;
using VisioAutomation.Shapes;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingDrawTests : VisioAutomationTest
    {
        [TestMethod]
        [DeploymentItem(@"datafiles\orgchart_1.xml", "datafiles")]
        public void Scripting_Draw_OrgChart()
        {
            // Load the chart
            string xml = this.get_datafile_content(@"datafiles\orgchart_1.xml");
            
            // Draw the Chart
            var client = this.GetScriptingClient();
            this.draw_org_chart(client, xml);
            
            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_DataTable()
        {
            var pagesize = new VisioAutomation.Drawing.Size(4, 4);
            var widths = new[] { 2.0, 1.5, 1.0 };
            double default_height = 0.25;
            var cellspacing = new VisioAutomation.Drawing.Size(0, 0);

            var items = new[]
                {
                    new {Name = "X", Age = 28, Score = 16},
                    new {Name = "Y", Age = 32, Score = 23},
                    new {Name = "Z", Age = 45, Score = 12},
                    new {Name = "U", Age = 48, Score = 10}
                };

            var dt = new DataTable();
            dt.Columns.Add("X", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Score", typeof(int));

            foreach (var item in items)
            {
                dt.Rows.Add(item.Name, item.Age, item.Score);
            }

            // Prepare the Page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);

            // Draw the table
            var heights = Enumerable.Repeat(default_height, items.Length).ToList();
            var shapes = client.Draw.Table(dt, widths, heights, cellspacing);

            // Verify
            int num_shapes_expected = items.Length*dt.Columns.Count;
            Assert.AreEqual(num_shapes_expected, shapes.Count);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_Grid()
        {
            var origin = new VisioAutomation.Drawing.Point(0, 4);
            var pagesize = new VisioAutomation.Drawing.Size(4, 4);
            var cellsize = new VisioAutomation.Drawing.Size(0.5, 0.25);
            int cols = 3;
            int rows = 6;

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);

            // Find the stencil and master
            var stencildoc = client.Document.OpenStencil("basic_u.vss");
            var master = client.Master.Get("Rectangle", stencildoc);

            // Draw the grid
            var page = client.Page.Get();
            var grid = new GridLayout(cols, rows, cellsize, master);
            grid.Origin = origin;
            grid.Render(page);

            // Verify
            int total_shapes_expected = cols*rows;
            var shapes = page.Shapes.ToEnumerable().ToList();
            int total_shapes_actual = shapes.Count;
            Assert.AreEqual(total_shapes_expected,total_shapes_actual);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_RectangleLineOval_0()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            var pagesize = new VA.Drawing.Size(4, 4);
            client.Page.New(pagesize, false);

            var shape_rect = client.Draw.Rectangle(1, 1, 3, 3);
            var shape_line = client.Draw.Line(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.Oval(0.2, 1, 3.8, 2);
            var shape_oval2 = client.Draw.Oval(new VA.Drawing.Point(2, 2), 0.5);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_BezierPolyLine_0()
        {
            var points = new[]
                {
                    new VA.Drawing.Point(0, 0),
                    new VA.Drawing.Point(2, 0.5),
                    new VA.Drawing.Point(2, 2),
                    new VA.Drawing.Point(3, 0.5)
                };
            var pagesize = new VA.Drawing.Size(4, 4);

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);
            
            // Draw the Shapes
            var shape_bezier = client.Draw.Bezier(points);
            var shape_polyline = client.Draw.PolyLine(points);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_PieSlice()
        {
            var pagesize = new VA.Drawing.Size(4, 4);
            var center = new VA.Drawing.Point(2, 2);
            double radius = 1.0;
            double start_angle = 0;
            double end_angle = Math.PI;

            // Create the page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);

            // Draw the Shape
            var shape = client.Draw.PieSlice(center, radius, start_angle, end_angle);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_PieChart()
        {
            var pagesize = new VA.Drawing.Size(4, 4);
            var center = new VA.Drawing.Point(2, 2);
            double radius = 1.0;
            
            var chart = new VisioAutomation.Models.Charting.PieChart(center, radius);
            chart.DataPoints.Add(1.0);
            chart.DataPoints.Add(2.0);
            chart.DataPoints.Add(3.0);
            chart.DataPoints.Add(4.0);

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);

            // Draw the chart

            client.Draw.PieChart(chart);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_BarChart()
        {
            var pagesize = new VA.Drawing.Size(4, 4);
            var rect1 = new VA.Drawing.Rectangle(0, 0, 4, 4);
            var rect2 = new VA.Drawing.Rectangle(5, 0, 9, 4);
            var rect3 = new VA.Drawing.Rectangle(10, 0, 14, 4);
            var bordersize = new VA.Drawing.Size(1.0, 1.0);

            var chart1 = new VisioAutomation.Models.Charting.BarChart(rect1);
            chart1.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(1.0));
            chart1.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(2.0));
            chart1.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(3.0));
            chart1.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(4.0));

            var chart2 = new VisioAutomation.Models.Charting.BarChart(rect2);
            chart2.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(1.0));
            chart2.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(2.0));
            chart2.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(-3.0));
            chart2.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(4.0));

            var chart3 = new VisioAutomation.Models.Charting.BarChart(rect3);
            chart3.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(-1.0));
            chart3.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(-2.0));
            chart3.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(-3.0));
            chart3.DataPoints.Add(new VisioAutomation.Models.Charting.DataPoint(-4.0));
            
            // Create the page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);

            // Draw the Charts
            client.Draw.BarChart(chart1);
            client.Draw.BarChart(chart2);
            client.Draw.BarChart(chart3);
            client.Page.ResizeToFitContents(bordersize,true);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_AreaChart()
        {
            var pagesize = new VA.Drawing.Size(4, 4);
            var rect1 = new VA.Drawing.Rectangle(0, 0, 4, 4);
            var rect2 = new VA.Drawing.Rectangle(5, 0, 9, 4);
            var rect3 = new VA.Drawing.Rectangle(10, 0, 14, 4);
            var padding = new VA.Drawing.Size(1.0, 1.0);
            
            var chart1 = new VisioAutomation.Models.Charting.AreaChart(rect1);
            chart1.DataPoints.Add(1.0);
            chart1.DataPoints.Add(2.0);
            chart1.DataPoints.Add(3.0);
            chart1.DataPoints.Add(4.0);

            var chart2 = new VisioAutomation.Models.Charting.AreaChart(rect2);
            chart2.DataPoints.Add(1.0);
            chart2.DataPoints.Add(2.0);
            chart2.DataPoints.Add(-3.0);
            chart2.DataPoints.Add(4.0);

            var chart3 = new VisioAutomation.Models.Charting.AreaChart(rect3);
            chart3.DataPoints.Add(-1.0);
            chart3.DataPoints.Add(-2.0);
            chart3.DataPoints.Add(-3.0);
            chart3.DataPoints.Add(-4.0);

            // Setup the Page
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(pagesize, false);

            // Draw the Charts
            client.Draw.AreaChart(chart1);
            client.Draw.AreaChart(chart2);
            client.Draw.AreaChart(chart3);
            client.Page.Get().ResizeToFitContents(padding);

            // Cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\directed_graph_1.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph1()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_1.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph1),".vsd");
            client.Document.SaveAs(output_filename);
            client.Document.Close(true);
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\directed_graph_2.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph2()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_2.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph2),".vsd");
            client.Document.SaveAs(output_filename);
            client.Document.Close(true);
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\directed_graph_3.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph3()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_3.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph3),".vsd");
            client.Document.SaveAs(output_filename);
            client.Document.Close(true);
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\directed_graph_4.xml", "datafiles")]
        public void Scripting_Draw_DirectedGraph4()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_4.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph4),".vsd");
            client.Document.SaveAs(output_filename);
            client.Document.Close(true);
        }

        public string get_datafile_content(string name)
        {
            string inputfilename = this.GetTestResultsOutPath( name );

            if (!File.Exists(inputfilename))
            {
                Assert.Fail("Could not locate " + inputfilename);
            }
            string text = File.ReadAllText(inputfilename);
            return text;
        }

        private void draw_directed_graph(VisioAutomation.Scripting.Client client, string dg_text)
        {
            var dg_xml = SXL.XDocument.Parse(dg_text);
            var dg_model = DirectedGraphBuilder.LoadFromXML(client, dg_xml);

            // TODO: Investigate if this this special case for Visio 2013 can be removed
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
            var pagesize = new VA.Drawing.Size(4, 4);
            var client = this.GetScriptingClient();
            
            // Create the page
            client.Document.New();
            client.Page.New(pagesize, false);

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencil("Basic_U.VSS");
            var master = client.Master.Get("Rectangle", basic_stencil);

            // Frop the Shapes
            client.Master.Drop(master, new VA.Drawing.Point(2, 2));

            // Verify
            var application = client.Application.Get();
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            Assert.AreEqual(1, shapes.Count);

            // cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Drop_Many()
        {
            var pagesize = new VA.Drawing.Size(10, 10);
            var client = this.GetScriptingClient();

            // Create the Page
            client.Document.New();
            client.Page.New(pagesize, false);

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencil("Basic_U.VSS");
            var m1 = client.Master.Get("Rectangle", basic_stencil);
            var m2 = client.Master.Get("Ellipse", basic_stencil);

            // Drop the Shapes
            var masters = new[] {m1, m2};
            var xys = new[] { 1.0, 2.0, 3.0, 4.0, 1.5, 4.5, 5.7, 2.4 };
            var points = VA.Drawing.Point.FromDoubles(xys).ToList();
            client.Master.Drop(masters, points);

            // Verify
            var application = client.Application.Get();
            Assert.AreEqual(4, application.ActivePage.Shapes.Count);

            // Cleanup
            client.Document.Close(true);
        }

        private void draw_org_chart(VisioAutomation.Scripting.Client client, string text)
        {
            var xmldoc = SXL.XDocument.Parse(text);
            var orgchart = OrgChartBuilder.LoadFromXml(client, xmldoc);
            client.Draw.OrgChart(orgchart);
        }

        [TestMethod]
        public void Scripting_Drop_Container_Master_Object()
        {
            var pagesize = new VA.Drawing.Size(4, 4);
            var client = this.GetScriptingClient();

            // Create the page
            client.Document.New();
            client.Page.New(pagesize, false);

            var application = client.Application.Get();
            var active_page = application.ActivePage;

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencil("Basic_U.VSS");
            var master = client.Master.Get("Rectangle", basic_stencil);

            // Drop the rectangle
            client.Master.Drop(master, new VA.Drawing.Point(2, 2) );

            // Select the rectangle... it should already be selected, but just make sure
            client.Selection.SelectAll();

            // Drop the container... since the rectangle is selected... it will automatically make it a member of the container
            var app = active_page.Application;

            var ver = client.Application.Version;
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";

            var stencil_type = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            var measurement_system = IVisio.VisMeasurementSystem.visMSUS;
            var containers_file = app.GetBuiltInStencilFile(stencil_type, measurement_system);
            var containers_doc = app.Documents.OpenStencil(containers_file);
            var masters = containers_doc.Masters;
            var container_master = masters.ItemU[cont_master_name];
            var dropped_container = client.Master.DropContainer(container_master);

            // Verify
            var shapes = active_page.Shapes;
            // There should be two shapes... the rectangle and the container
            Assert.AreEqual(2, shapes.Count);

            // Verify that we did indeed drop a container
            Assert.AreEqual("Container",
                UserDefinedCellHelper
                    .Get(dropped_container)
                    .First(s => s.Name == "msvStructureType")
                    .Value.Result);

            // cleanup
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Drop_Container_Master_Name()
        {
            var pagesize = new VA.Drawing.Size(4, 4);
            var client = this.GetScriptingClient();

            // Create the page
            client.Document.New();
            client.Page.New(pagesize, false);

            var application = client.Application.Get();
            var active_page = application.ActivePage;

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencil("Basic_U.VSS");
            var master = client.Master.Get("Rectangle", basic_stencil);

            // Drop the rectangle
            client.Master.Drop(master, new VA.Drawing.Point(2, 2) );

            // Select the rectangle... it should already be selected, but just make sure
            client.Selection.SelectAll();

            // Drop the container... since the rectangle is selected... it will automatically make it a member of the container
            var ver = client.Application.Version;
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var dropped_container = client.Master.DropContainer(cont_master_name);

            // Verify
            var shapes = active_page.Shapes;
            // There should be two shapes... the rectangle and the container
            Assert.AreEqual(2, shapes.Count);

            // Verify that we did indeed drop a container
            Assert.AreEqual("Container",
                UserDefinedCellHelper
                    .Get(dropped_container)
                    .First(s => s.Name == "msvStructureType")
                    .Value.Result);

            // cleanup
            client.Document.Close(true);
        }
    }
}