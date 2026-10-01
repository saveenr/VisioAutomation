using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAORGCHART = VisioAutomation.Models.Documents.OrgCharts;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using SXL = System.Xml.Linq;

namespace VTest.Models
{
    [MUT.TestClass]
    public class OrgChartDrawModelTests: Framework.VTest
    {
        [MUT.TestMethod]
        public void RenderOrgChart_FromSampleData_DoesNotThrow()
        {
            // Load the chart
            string xml = this.get_datafile_content(@"datafiles\orgchart_1.xml");
            
            // Draw the Chart
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            this.draw_org_chart(client, xml);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawOrgChart_DrawsInANewDocumentAndLeavesTheTargetPageAlone()
        {
            var orgchart = new VAORGCHART.OrgChartDocument();
            orgchart.OrgCharts.Add(new VAORGCHART.Node("A"));

            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var app = this.GetVisioApplication();
            var target_doc = app.ActiveDocument;
            var target_page = app.ActivePage;
            target_page.DrawRectangle(new VA.Core.Rectangle(1, 1, 2, 2));
            double width_before = target_page.PageSheet.CellsU["PageWidth"].ResultIU;
            double height_before = target_page.PageSheet.CellsU["PageHeight"].ResultIU;

            client.Model.DrawOrgChart(VisioScripting.TargetPage.Auto, orgchart);

            MUT.Assert.AreNotEqual(target_doc.ID, app.ActiveDocument.ID, "the chart should be drawn in a new document");
            MUT.Assert.AreEqual(width_before, target_page.PageSheet.CellsU["PageWidth"].ResultIU, 1e-9, "the target page width should not change");
            MUT.Assert.AreEqual(height_before, target_page.PageSheet.CellsU["PageHeight"].ResultIU, 1e-9, "the target page height should not change");

            app.ActiveDocument.Close(true);
            target_doc.Close(true);
        }

        private void draw_org_chart(VisioScripting.Client client, string text)
        {
            var xmldoc = SXL.XDocument.Parse(text);
            var orgchart = client.Model.LoadOrgChartFromXml(xmldoc);

            client.Model.DrawOrgChart(VisioScripting.TargetPage.Auto, orgchart);
        }

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

            var app = this.GetVisioApplication();
            orgchart.Render(app);

            var active_page = app.ActivePage;
            var page = active_page;
            page.ResizeToFitContents();

            app.ActiveDocument.Close(true);
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

            var app = this.GetVisioApplication();

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
            MUT.Assert.AreEqual(orgchart_doc.OrgChartLayoutOptions.DefaultNodeSize, Framework.VTest.GetSize(n_b.VisioShape));

            app.ActiveDocument.Close(true);
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

            var app = this.GetVisioApplication();

            orgchart.Render(app);

            app.ActiveDocument.Close(true);
        }

        [MUT.TestMethod]
        public void OrgChart_MultipleOrgCharts_EachPageShowsItsOwnRoot()
        {
            var orgchart = new VAORGCHART.OrgChartDocument();

            var n_a = new VAORGCHART.Node("A");
            var n_b = new VAORGCHART.Node("B");
            n_a.Children.Add(n_b);

            var n_x = new VAORGCHART.Node("X");
            var n_y = new VAORGCHART.Node("Y");
            n_x.Children.Add(n_y);

            orgchart.OrgCharts.Add(n_a);
            orgchart.OrgCharts.Add(n_x);

            var app = this.GetVisioApplication();
            orgchart.Render(app);
            var doc = app.ActiveDocument;

            MUT.Assert.AreEqual(2, doc.Pages.Count);

            var page1_texts = doc.Pages[1].Shapes.Cast<IVisio.Shape>().Select(s => s.Text).ToList();
            var page2_texts = doc.Pages[2].Shapes.Cast<IVisio.Shape>().Select(s => s.Text).ToList();

            MUT.Assert.IsTrue(page1_texts.Contains("A") && page1_texts.Contains("B"), "page 1 should show the first chart");
            MUT.Assert.IsFalse(page1_texts.Contains("X") || page1_texts.Contains("Y"), "page 1 should not show the second chart");
            MUT.Assert.IsTrue(page2_texts.Contains("X") && page2_texts.Contains("Y"), "page 2 should show the second chart");
            MUT.Assert.IsFalse(page2_texts.Contains("A") || page2_texts.Contains("B"), "page 2 should not show the first chart");

            doc.Close(true);
        }
    }
}