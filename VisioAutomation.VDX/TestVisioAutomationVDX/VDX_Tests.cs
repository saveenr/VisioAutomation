using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace TestVisioAutomationVDX
{
    [TestClass]
    public class VDX_Tests : TestVisioAutomation.VisioAutomationTest
    {
        public IVisio.Document TryOpen(IVisio.Documents docs, string filename)
        {
            using (var scope = new VA.Application.AlertResponseScope(docs.Application,VA.Application.AlertResponseCode.No))
            {
                var doc = docs.Open(filename);
                return doc;
            }
        }

        public void VerifyDocCanBeLoaded(string filename)
        {
            var app = new IVisio.Application();
            var version = VA.Application.ApplicationHelper.GetVersion(app);
            string logfilename = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);

            VA.Application.Logging.XmlErrorLog log_before=null;
            var old_fileinfo = new System.IO.FileInfo(logfilename);

            if (System.IO.File.Exists(logfilename))
            {
                log_before = new VA.Application.Logging.XmlErrorLog(logfilename);
            }

            var time = System.DateTime.Now;
            this.TryOpen(app.Documents, filename); // this causes the doc to load no matter what the error 

            VA.Application.Logging.XmlErrorLog log_after = null;
            if (System.IO.File.Exists(logfilename))
            {
                log_after = new VA.Application.Logging.XmlErrorLog(logfilename);
            }

            if (log_before != null && log_after == null)
            {
                Assert.Fail("Invalid case for all visio versions - if it existed before it must exist after");
                return;
            }
 
            if (log_before == null && log_after == null)
            {
                // Didn't exist before, didn't exist after - that's fine - the file loaded with no issues
                return;
            }

            // log_after exists
            VDX_Tests.VerifyNoErrorsInLog(log_after, filename, logfilename, version, time);

            VA.Documents.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit();
        }

        private static void VerifyNoErrorsInLog(VA.Application.Logging.XmlErrorLog log_after, string filename, string logfilename, System.Version version, System.DateTime opentime)
        {
            int duration = 2;
            var lower_time_bound = opentime;
            var upper_time_bound = opentime.AddSeconds(duration);

            // First see of any sessions matching the source exist
            // If not, then the load must have been successful
            var all_sessions_from_source = log_after.FileSessions.Where(s => s.Source == filename).ToList();
            if (all_sessions_from_source.Count < 1)
            {
                return;
            }

            // From that set of sessions, find the one closest in time to when we
            // asked Visio to open the file.
            // If none could be found, then then we assume no error could be found
            // NOTE: It would be better not to have to rely on a time duration.
            var sessions_near_in_time =
                all_sessions_from_source.Where(c => (lower_time_bound <= c.StartTime && c.StartTime <= upper_time_bound)).ToList();
            if (sessions_near_in_time.Count < 1)
            {
                // couldn't find any in the time period. We'll have to assume no error could be found.
                return;
            }

            var target_session = all_sessions_from_source[0];

            var warnings = target_session.Records.Where(rec => rec.Type == "Warning").ToList();
            var errors = target_session.Records.Where(rec => rec.Type == "Error").ToList();

            if (errors.Count > 0)
            {                
                string msg = string.Format("XML Error Log {0} contains an error", logfilename);
                Assert.Fail(msg);
            }

        }

        [TestMethod]
        public void VDX_MultiPageDocument()
        {
            string output_filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");

            var template = new VA.VDX.Template(); // the default template
            var doc = new VA.VDX.Elements.Drawing(template);

            var Page01 = VDX_Files.GetPage01_Simple_Fill_Format(doc);
            var Page02 = VDX_Files.GetPage02_Locking(doc);
            var Page03 = VDX_Files.GetPage03_Text_Block(doc);
            var Page04 = VDX_Files.GetPage04_Simple_Text(doc);
            var Page05 = VDX_Files.GetPage05_Formatted_Text(doc);
            var Page06 = VDX_Files.GetPage06_All_FillPatterns(doc);
            var Page08 = VDX_Files.GetPage08_Connector_With_Geometry(doc);
            var Page09 = VDX_Files.GetPage09_Layout(doc);
            var Page10 = VDX_Files.GetPage10_layers(doc);
            var Page11 = VDX_Files.GetPage11_Add_color(doc);
            var Page12 = VDX_Files.GetPage12_AdjustToTextSize(doc);
            var Page13 = VDX_Files.GetPage13_MultipleConnectors(doc);
            var Page14 = VDX_Files.GetPage14_Hyperlinks(doc);

            var w1 = new VA.VDX.Elements.DocumentWindow();
            w1.ShowGrid = false;
            w1.ShowGuides = false;
            w1.ShowConnectionPoints = false;
            w1.ShowPageBreaks = false;
            w1.Page = Page01.ID; // point to first page we created
            
            doc.Windows.Add(w1);

            doc.Save(output_filename);
            
            // Verify this file can be loaded
            this.VerifyDocCanBeLoaded(output_filename);
        }

        [TestMethod]
        public void VDX_CustomProperties()
        {
            string filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");

            var template = new VA.VDX.Template();
            var doc_node = new VA.VDX.Elements.Drawing(template);

            int rect_id = doc_node.GetMasterMetaData("REctAngle").ID;

            var node_page = new VA.VDX.Elements.Page(8, 5);
            doc_node.Pages.Add(node_page);

            var node_shape = new VA.VDX.Elements.Shape(rect_id, 4, 2, 3, 2);
            node_shape.CustomProps = new VA.VDX.Elements.CustomProps();

            var node_custprop0 = new VA.VDX.Elements.CustomProp("PROP1");
            node_custprop0.Value = "VALUE1";
            node_shape.CustomProps.Add(node_custprop0);

            var node_custprop1 = new VA.VDX.Elements.CustomProp("PROP2");
            node_custprop1.Value = "123";
            node_custprop1.Type.Result = VisioAutomation.VDX.Enums.CustomPropType.String;
            node_shape.CustomProps.Add(node_custprop1);

            var node_custprop2 = new VA.VDX.Elements.CustomProp("PROP3");
            node_custprop2.Value = "456";
            node_custprop2.Type.Result = VisioAutomation.VDX.Enums.CustomPropType.Number;
            node_shape.CustomProps.Add(node_custprop2);

            node_page.Shapes.Add(node_shape);

            doc_node.Save(filename);

            var app = new IVisio.Application();
            var docs = app.Documents;
            var doc = docs.Add(filename);

            var page = app.ActivePage;
            var shapes = page.Shapes;
            Assert.AreEqual(1,page.Shapes.Count);

            var shape = page.Shapes[1];
            var customprops = VA.Shapes.CustomProperties.CustomPropertyHelper.Get(shape);

            Assert.IsTrue(customprops.ContainsKey("PROP1"));
            Assert.AreEqual("\"VALUE1\"",customprops["PROP1"].Value.Formula);


            Assert.IsTrue(customprops.ContainsKey("PROP2"));
            Assert.AreEqual("\"123\"", customprops["PROP2"].Value.Formula);
            Assert.AreEqual("0", customprops["PROP2"].Type.Formula);

            Assert.IsTrue(customprops.ContainsKey("PROP3"));
            Assert.AreEqual("\"456\"", customprops["PROP3"].Value.Formula);
            Assert.AreEqual("2", customprops["PROP3"].Type.Formula);

            app.Quit(true);
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\template_router.vdx", "datafiles")]
        public void VDX_CustomTemplate()
        {
            string input_filename = this.GetTestResultsOutPath(@"datafiles\template_router.vdx");
            string output_filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");
            
            // Load the template
            string template_xml = System.IO.File.ReadAllText(input_filename);

            var template = new VA.VDX.Template(template_xml);
            var doc = new VisioAutomation.VDX.Elements.Drawing(template);
            var page = new VA.VDX.Elements.Page(8, 4);

            doc.Pages.Add(page);

            // add layers
            var layer0 = page.AddLayer("Layer0", 0);
            var layer1 = page.AddLayer("Layer1", 1);
            var layer2 = page.AddLayer("Layer2", 2);

            // create layout
            var layout = new VA.VDX.Sections.Layout();
            layout.ShapeRouteStyle.Result = VA.VDX.Enums.RouteStyle.TreeEW;

            // find the id of the master for rounded rectangles
            int shapeMasterNameId = doc.GetMasterMetaData("Router").ID;
            bool shapeMasterNameGroup = doc.GetMasterMetaData("Router").IsGroup;

            // add shape1
            var shape1 = new VA.VDX.Elements.Shape(shapeMasterNameId, shapeMasterNameGroup, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Router1");
            shape1.Layout = layout;

            // add shape2
            var shape2 = new VA.VDX.Elements.Shape(shapeMasterNameId, shapeMasterNameGroup, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Router2");
            shape2.Layout = shape1.Layout;

            // add shape3 - this is the dynamic connector
            VA.VDX.Elements.Shape shape3 = VA.VDX.Elements.Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.BeginX.Result = 1;
            shape3.XForm1D.EndX.Result = 5;
            shape3.XForm1D.BeginY.Result = 3;
            shape3.XForm1D.EndY.Result = 3;
            page.Shapes.Add(shape3);
            shape3.Geom = new VA.VDX.Sections.Geom();
            shape3.Geom.Rows.Add(new VA.VDX.Sections.MoveTo(1, 3));
            shape3.Geom.Rows.Add(new VA.VDX.Sections.LineTo(5, 3));

            shape3.Layout = shape1.Layout;
            page.ConnectShapesViaConnector(shape3, shape1, shape2);

            // handle layers
            shape3.LayerMembership = new List<int> {layer0.Index, layer2.Index};
            shape1.LayerMembership = new List<int> {layer1.Index};
            shape2.LayerMembership = new List<int> {layer2.Index};

            // write document to disk as .vdx file

            doc.Save(output_filename);

            this.VerifyDocCanBeLoaded(output_filename);
        }

        [TestMethod]
        [DeploymentItem(@"datafiles\template_router.vdx", "datafiles")]
        public void VDX_CheckNoErrorOnLoad()
        {
            var folder = this.TestResultsOutFolder;
            
            string input_filename = this.GetTestResultsOutPath(@"datafiles\template_router.vdx");

            this.VerifyDocCanBeLoaded(input_filename);
        }


        [TestMethod]
        [DeploymentItem(@"datafiles\vdx_with_warnings_1.vdx", "datafiles")]
        public void VDX_DetectLoadWarnings()
        {
            string input_filename = this.GetTestResultsOutPath(@"datafiles\vdx_with_warnings_1.vdx");
 
            // Load the VDX
            var app = new IVisio.Application();
            var version = VA.Application.ApplicationHelper.GetVersion(app);
            string logfilename = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);

            var doc = this.TryOpen(app.Documents, input_filename);
            
            // See what happened
            var log_after = new VA.Application.Logging.XmlErrorLog(logfilename);
            var most_recent_session = log_after.FileSessions[0];
            var warnings = most_recent_session.Records.Where(r => r.Type == "Warning").ToList();
            var errors = most_recent_session.Records.Where(r => r.Type == "Error").ToList();

            // Verify
            int expected_errors = 0;  // this VDX should not report any errors
            int expected_warnings = 4; // this VDX should contain four warnings for Visio2010 and two warnings for Visio 2013         
            if (version.Major >= 15)
            {
                expected_warnings = 2;
            }

            Assert.AreEqual(expected_errors, errors.Count); // this VDX should not report any errors
            Assert.AreEqual(expected_warnings, warnings.Count); // this VDX should contain exactly two warnings                                
            Assert.AreEqual(1, app.Documents.Count);

            // Cleanup
            VA.Documents.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit(true);
        }
    }
}