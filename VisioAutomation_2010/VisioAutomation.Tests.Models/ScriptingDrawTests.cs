using System.Data;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioScripting;
using GRID = VisioAutomation.Models.Layouts.Grid;
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
            client.Document.NewDocument();
            this.draw_org_chart(client, xml);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_Draw_DataTable()
        {
            var pagesize = new VisioAutomation.Geometry.Size(4, 4);
            var widths = new[] { 2.0, 1.5, 1.0 };
            double default_height = 0.25;
            var cellspacing = new VisioAutomation.Geometry.Size(0, 0);

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
            client.Document.NewDocument();

            var page = client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Draw the table
            var heights = Enumerable.Repeat(default_height, items.Length).ToList();

            var shapes = client.Model.DrawDataTable(VisioScripting.TargetPage.Auto, dt, widths, heights, cellspacing);

            // Verify
            int num_shapes_expected = items.Length*dt.Columns.Count;
            Assert.AreEqual(num_shapes_expected, shapes.Count);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_Draw_Grid()
        {
            var origin = new VisioAutomation.Geometry.Point(0, 4);
            var pagesize = new VisioAutomation.Geometry.Size(4, 4);
            var cellsize = new VisioAutomation.Geometry.Size(0.5, 0.25);
            int cols = 3;
            int rows = 6;

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Find the stencil and master
            var stencildoc = client.Document.OpenStencilDocument("basic_u.vss");
            var stencil_targetdoc = new VisioScripting.TargetDocument(stencildoc);
            var master = client.Master.GetMaster(stencil_targetdoc, "Rectangle");

            // Draw the grid
            var page = client.Page.GetActivePage();
            var grid = new GRID.GridLayout(cols, rows, cellsize, master);
            grid.Origin = origin;
            grid.Render(page);

            // Verify
            int total_shapes_expected = cols*rows;
            var shapes = page.Shapes.ToList();
            int total_shapes_actual = shapes.Count;
            Assert.AreEqual(total_shapes_expected,total_shapes_actual);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_Draw_RectangleLineOval_0()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var pagesize = new VA.Geometry.Size(4, 4);

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var shape_rect = client.Draw.DrawRectangle(1, 1, 3, 3);
            var shape_line = client.Draw.DrawLine(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.DrawOval(0.2, 1, 3.8, 2);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_Draw_BezierPolyLine_0()
        {
            var points = new[]
                {
                    new VA.Geometry.Point(0, 0),
                    new VA.Geometry.Point(2, 0.5),
                    new VA.Geometry.Point(2, 2),
                    new VA.Geometry.Point(3, 0.5)
                };
            var pagesize = new VA.Geometry.Size(4, 4);

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);
            
            // Draw the Shapes
            var shape_bezier = client.Draw.DrawBezier(points);
            var shape_polyline = client.Draw.DrawPolyLine(points);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
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

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
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
            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
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

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
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

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        public string get_datafile_content(string name)
        {
            string inputfilename = this._get_test_results_out_path( name );

            if (!File.Exists(inputfilename))
            {
                Assert.Fail("Could not locate " + inputfilename);
            }
            string text = File.ReadAllText(inputfilename);
            return text;
        }

        private void draw_directed_graph(VisioScripting.Client client, string dg_text)
        {
            var dg_xml = SXL.XDocument.Parse(dg_text);
            var dg_model = VisioScripting.Builders.DirectedGraphBuilder.LoadFromXml(client, dg_xml);

            // TODO: Investigate if this this special case for Visio 2013 can be removed
            // this is a temporary fix to handle the fact that server_u.vss in Visio 2013 doesn't result in server_u.vssx 
            // gettign automatically loaded

            var version = client.Application.ApplicationVersion;
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

            client.Model.DrawDirectedGraphDocument(dg_model);
        }
        
        [TestMethod]
        public void Scripting_Drop_Master()
        {
            var pagesize = new VA.Geometry.Size(4, 4);
            var client = this.GetScriptingClient();

            // Create the page
            client.Document.NewDocument();

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var master = client.Master.GetMaster(stencil_targetdoc, "Rectangle");

            // Frop the Shapes

            client.Master.DropMaster(VisioScripting.TargetPage.Auto, master, new VA.Geometry.Point(2, 2));

            // Verify
            var application = client.Application.GetAttachedApplication();
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            Assert.AreEqual(1, shapes.Count);

            // cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_Drop_Many()
        {
            var pagesize = new VA.Geometry.Size(10, 10);
            var client = this.GetScriptingClient();

            // Create the Page
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var m1 = client.Master.GetMaster(stencil_targetdoc, "Rectangle");
            var m2 = client.Master.GetMaster(stencil_targetdoc, "Ellipse");

            // Drop the Shapes
            var masters = new[] {m1, m2};
            var xys = new[] { 1.0, 2.0, 3.0, 4.0, 1.5, 4.5, 5.7, 2.4 };
            var points = VA.Geometry.Point.FromDoubles(xys).ToList();

            client.Master.DropMasters(VisioScripting.TargetPage.Auto, masters, points);

            // Verify
            var application = client.Application.GetAttachedApplication();
            Assert.AreEqual(4, application.ActivePage.Shapes.Count);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        private void draw_org_chart(VisioScripting.Client client, string text)
        {
            var xmldoc = SXL.XDocument.Parse(text);
            var orgchart = VisioScripting.Builders.OrgChartBuilder.LoadFromXml(client, xmldoc);

            client.Model.DrawOrgChart(VisioScripting.TargetPage.Auto, orgchart);
        }

        [TestMethod]
        public void Scripting_Drop_Container_Master_Object()
        {
            var pagesize = new VA.Geometry.Size(4, 4);
            var client = this.GetScriptingClient();


            // Create the page
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var application = client.Application.GetAttachedApplication();
            var active_page = application.ActivePage;

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var master = client.Master.GetMaster(stencil_targetdoc, "Rectangle");

            // Drop the rectangle
            client.Master.DropMaster(VisioScripting.TargetPage.Auto, master, new VA.Geometry.Point(2, 2) );

            // Select the rectangle... it should already be selected, but just make sure


            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            // Drop the container... since the rectangle is selected... it will automatically make it a member of the container
            var app = active_page.Application;

            var ver = client.Application.ApplicationVersion;
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";

            var stencil_type = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            var measurement_system = IVisio.VisMeasurementSystem.visMSUS;
            var containers_file = app.GetBuiltInStencilFile(stencil_type, measurement_system);
            var containers_doc = app.Documents.OpenStencil(containers_file);
            var masters = containers_doc.Masters;
            var container_master = masters.ItemU[cont_master_name];

            var dropped_container = client.Container.DropContainerMaster(VisioScripting.TargetPage.Auto, container_master);

            // Verify
            var shapes = active_page.Shapes;
            // There should be two shapes... the rectangle and the container
            Assert.AreEqual(2, shapes.Count);

            // Verify that we did indeed drop a container

            var results_dic = VisioAutomation.Shapes.UserDefinedCellHelper.GetDictionary(dropped_container, VA.ShapeSheet.CellValueType.Result);
            Assert.IsTrue(results_dic.ContainsKey("msvStructureType"));
            var prop = results_dic["msvStructureType"];
            Assert.AreEqual("Container", prop.Value.Value);

            // cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_Drop_Container_Master_Name()
        {
            var pagesize = new VA.Geometry.Size(4, 4);
            var client = this.GetScriptingClient();

            // Create the page
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var application = client.Application.GetAttachedApplication();
            var active_page = application.ActivePage;

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var basic_stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var master = client.Master.GetMaster(basic_stencil_targetdoc, "Rectangle");

            // Drop the rectangle
            client.Master.DropMaster(VisioScripting.TargetPage.Auto, master, new VA.Geometry.Point(2, 2) );


            // Select the rectangle... it should already be selected, but just make sure
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            // Drop the container... since the rectangle is selected... it will automatically make it a member of the container
            var ver = client.Application.ApplicationVersion;
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var dropped_container = client.Container.DropContainer(VisioScripting.TargetPage.Auto, cont_master_name);

            // Verify
            var shapes = active_page.Shapes;
            // There should be two shapes... the rectangle and the container
            Assert.AreEqual(2, shapes.Count);

            // Verify that we did indeed drop a container           
            var results_dic = VisioAutomation.Shapes.UserDefinedCellHelper.GetDictionary(dropped_container, VA.ShapeSheet.CellValueType.Result);
            Assert.IsTrue(results_dic.ContainsKey("msvStructureType"));
            var prop = results_dic["msvStructureType"];
            Assert.AreEqual("Container", prop.Value.Value);

            // cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }
    }
}