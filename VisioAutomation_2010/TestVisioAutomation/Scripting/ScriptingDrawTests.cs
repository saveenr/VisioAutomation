using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingDrawTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Draw_DataTable_0()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var items = new[]
                            {
                                new {Name = "X", Age = 28, Score = 16},
                                new {Name = "Y", Age = 32, Score = 23},
                                new {Name = "Z", Age = 45, Score = 12},
                                new {Name = "U", Age = 48, Score = 10}
                            };

            var dt = new System.Data.DataTable();
            dt.Columns.Add("X", typeof (string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Score", typeof(int));

            foreach (var item in items)
            {
                dt.Rows.Add(item.Name, item.Age, item.Score);
            }

            var widths = new[] {2.0, 1.5, 1.0};
            var heights = Enumerable.Repeat(0.25, items.Length).ToList();
            var shapes = ss.Draw.Table(dt, widths, heights, new VA.Drawing.Size(0, 0));

            Assert.AreEqual(items.Length*3, shapes.Count);

            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_Grid()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var cellsize = new VA.Drawing.Size(0.5, 0.25);
            int cols = 3;
            int rows = 6;

            ss.Document.OpenStencil("basic_u.vss");
            string stencil = "basic_u.vss";
            string mastername = "Rectangle";
            var master = ss.Master.Get(mastername, stencil);

            var grid = new VA.Layout.Models.Grid.GridLayout(cols, rows, cellsize, master);
            grid.Origin = new VA.Drawing.Point(0, 4);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_RectangleLineOval_0()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var shape_rect = ss.Draw.Rectangle(1, 1, 3, 3);
            var shape_line = ss.Draw.Line(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = ss.Draw.Oval(0.2, 1, 3.8, 2);
            var shape_oval2 = ss.Draw.Oval(new VA.Drawing.Point(2, 2), 0.5);

            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_BezierPolyLine_0()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var points = new[]
                             {
                                 new VA.Drawing.Point(0, 0),
                                 new VA.Drawing.Point(2, 0.5),
                                 new VA.Drawing.Point(2, 2),
                                 new VA.Drawing.Point(3, 0.5),
                             };

            var shape_bezier = ss.Draw.Bezier(points);
            var shape_polyline = ss.Draw.PolyLine(points);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_PieSlice()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var center = new VA.Drawing.Point(2, 2);
            double radius = 1.0;
            double start_angle = 0;
            double end_angle = System.Math.PI;

            var shape = ss.Draw.PieSlice(center, radius, start_angle, end_angle);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Draw_PieSlices()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var center = new VA.Drawing.Point(2, 2);
            double radius = 1.0;
            double[] values = new[] {1.0, 2.0, 3.0, 4.0};
            var shapes = ss.Draw.PieSlices(center, radius, values);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_DirectedGraph1()
        {
            var ss = GetScriptingSession();
            draw_dg(ss, TestVisioAutomation.Properties.Resources.sampleflowchart1);
            ss.Document.Close(true);
            draw_dg(ss, TestVisioAutomation.Properties.Resources.sampleflowchart3);
            ss.Document.Close(true);
            draw_dg(ss, TestVisioAutomation.Properties.Resources.sampleflowchart4);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_DirectedGraph4()
        {
            var ss = GetScriptingSession();
            draw_dg(ss, TestVisioAutomation.Properties.Resources.sampleflowchart4);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_DirectedGraph3()
        {
            var ss = GetScriptingSession();
            draw_dg(ss, TestVisioAutomation.Properties.Resources.sampleflowchart3);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_DirectedGraph2()
        {
            var ss = GetScriptingSession();
            draw_dg(ss, TestVisioAutomation.Properties.Resources.sampleflowchart2);
            ss.Document.Close(true);
        }


        private void draw_dg(VA.Scripting.Session scriptingsession, string dg_text)
        {
            var dg_xml = System.Xml.Linq.XDocument.Parse(dg_text);
            var dg_model = VA.Scripting.DirectedGraph.DirectedGraphBuilder.LoadFromXML(scriptingsession, dg_xml);
            VA.Scripting.DirectedGraph.DirectedGraphBuilder.RenderDiagrams(scriptingsession, dg_model);
        }

        [TestMethod]
        public void Scripting_Orgchart()
        {
            var ss = GetScriptingSession();
            draw_org_chart(ss, TestVisioAutomation.Properties.Resources.sampleorgchart1);
            ss.Document.Close(true);
            VA.Documents.DocumentHelper.ForceCloseAll(ss.VisioApplication.Documents);
        }

        private void draw_org_chart(VA.Scripting.Session scriptingsession, string text)
        {
            var xmldoc = System.Xml.Linq.XDocument.Parse(text);
            var orgchart = VA.Scripting.OrgChart.OrgChartBuilder.LoadFromXML(scriptingsession, xmldoc);
            scriptingsession.Draw.OrgChart(orgchart);
        }


        [TestMethod]
        public void Scripting_DropMaster()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);
            ss.Document.OpenStencil("Basic_U.VSS");
            var master = ss.Master.Get("Rectangle", "Basic_U.VSS");
            ss.Master.Drop(master, 2, 2);
            var application = ss.VisioApplication;
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            Assert.AreEqual(1, shapes.Count);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_DropMany()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(10, 10), false);
            ss.Document.OpenStencil("Basic_U.VSS");

            var m1 = ss.Master.Get("Rectangle", "Basic_U.VSS");
            var m2 = ss.Master.Get("Ellipse", "Basic_U.VSS");

            var masters = new[] {m1, m2};
            var points = VA.Drawing.Point.FromDoubles(new[] { 1.0, 2.0, 3.0, 4.0, 1.5, 4.5, 5.7, 2.4 }).ToList();

            ss.Master.Drop(masters, points);
            
            Assert.AreEqual(4, ss.VisioApplication.ActivePage.Shapes.Count);
            ss.Document.Close(true);
        }
    }
}