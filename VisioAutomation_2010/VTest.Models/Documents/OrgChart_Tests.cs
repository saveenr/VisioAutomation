using System.Linq;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAORGCHART = VisioAutomation.Models.Documents.OrgCharts;

namespace VTest.Models.Documents
{
    [MUT.TestClass]
    public class OrgChart_Tests : Framework.VTest
    {
        [MUT.TestMethod]
        public void OrgChart_MustHaveContent()
        {
            // Before an Org Chart is rendered it must have at least one node
            bool caught = false;
            var orgcgart = new VAORGCHART.OrgChartDocument();
            var page1 = this.GetNewPage(this.StandardPageSize);
            try
            {
                var application = page1.Application;
                orgcgart.Render(application);
            }
            catch (System.Exception)
            {
                page1.Delete(0);
                caught = true;
            }
            if (caught == false)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }
        }

        [MUT.TestMethod]
        public void OrgChart_SingleNode()
        {
            // Draw the minimum org chart - a chart with one nod
            var orgchart = new VAORGCHART.OrgChartDocument();

            var n_a = new VAORGCHART.Node("A");
            n_a.Size = new VA.Core.Size(4, 2);
            orgchart.OrgCharts.Add(n_a);

            var app = new IVisio.Application();
            orgchart.Render(app);

            var active_page = app.ActivePage;
            var page = active_page;
            page.ResizeToFitContents();

            app.Quit(true);
        }

        [MUT.TestMethod]
        public void OrgChart_FiveNodes()
        {
            // Verify that basic org chart connectivity is maintained

            var orgchart_doc = new VAORGCHART.OrgChartDocument();

            var n_a = new VAORGCHART.Node("A");
            var n_b = new VAORGCHART.Node("B");
            var n_c = new VAORGCHART.Node("C");
            var n_d = new VAORGCHART.Node("D");
            var n_e = new VAORGCHART.Node("E");

            n_a.Children.Add(n_b);
            n_a.Children.Add(n_c);
            n_c.Children.Add(n_d);
            n_c.Children.Add(n_e);

            n_a.Size = new VA.Core.Size(4, 2);

            orgchart_doc.OrgCharts.Add(n_a);

            var app = new IVisio.Application();
            
            orgchart_doc.Render(app);

            var active_page = app.ActivePage;
            var page = active_page;
            page.ResizeToFitContents();

            var shapes = active_page.Shapes.ToList();
            var shapes_2d = shapes.Where(s => s.OneD == 0).ToList();
            var shapes_1d = shapes.Where(s => s.OneD != 0).ToList();
            var shapes_connector = shapes.Where(s => s.Master.NameU == "Dynamic connector").ToList();

            MUT.Assert.AreEqual(5 + 4, shapes.Count);
            MUT.Assert.AreEqual(5, shapes_2d.Count);
            MUT.Assert.AreEqual(4, shapes_1d.Count);
            MUT.Assert.AreEqual(4, shapes_connector.Count);

            MUT.Assert.AreEqual("A", n_a.VisioShape.Text.Trim());
                // trimming because extra ending space is added (don't know why)
            MUT.Assert.AreEqual("B", n_b.VisioShape.Text.Trim());
            MUT.Assert.AreEqual("C", n_c.VisioShape.Text.Trim());
            MUT.Assert.AreEqual("D", n_d.VisioShape.Text.Trim());
            MUT.Assert.AreEqual("E", n_e.VisioShape.Text.Trim());

            MUT.Assert.AreEqual(new VA.Core.Size(4, 2), Framework.VTest.GetSize(n_a.VisioShape));
            MUT.Assert.AreEqual(orgchart_doc.OrgChartLayoutOptions.DefaultNodeSize,  Framework.VTest.GetSize(n_b.VisioShape));

            app.Quit(true);
        }

        [MUT.TestMethod]
        public void OrgChart_MultipleOrgCharts()
        {
            // Verify that we can create multiple org charts in one
            // document

            var orgchart = new VAORGCHART.OrgChartDocument();

            var n_a = new VAORGCHART.Node("A");
            var n_b = new VAORGCHART.Node("B");
            var n_c = new VAORGCHART.Node("C");
            var n_d = new VAORGCHART.Node("D");
            var n_e = new VAORGCHART.Node("E");

            n_a.Children.Add(n_b);
            n_a.Children.Add(n_c);
            n_c.Children.Add(n_d);
            n_c.Children.Add(n_e);

            n_a.Size = new VA.Core.Size(4, 2);

            orgchart.OrgCharts.Add(n_a);
            orgchart.OrgCharts.Add(n_a);

            var app = new IVisio.Application();

            orgchart.Render(app);

            app.Quit(true);
        }
    }
}