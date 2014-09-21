using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace TestVisioAutomationVDX
{
    [TestClass]
    public class VDX_Tests
    {
        public IVisio.Document TryOpen(IVisio.Documents docs, string filename)
        {
            var app = docs.Application;
            using (var scope = new VA.Application.AlertResponseScope(app,VA.Application.AlertResponseCode.No))
            {
                var doc = app.Documents.Open(filename);
                return doc;
            }
        }

        public void CheckIfLoadsWithoutErrorLog(string filename)
        {
            var app = new IVisio.Application();

            DeleteXmlErrorLog(app);

            // this causes the doc to load no matter what the error ))))))
            var doc = TryOpen(app.Documents, filename);

            if (XmlErrorLogExists(app))
            {
                string logfilename = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);
                bool exists = System.IO.File.Exists(logfilename);

                var log = new VA.Application.Logging.LogFile(logfilename);

                var session = log.Sessions[log.Sessions.Count - 1];
                foreach (var rec in session.Records)
                {
                    if (rec.Type != "Warning")
                    {
                        string msg = string.Format("XML Error Log {0} Was Created when opening the VDX file", logfilename);
                        Assert.Fail(msg);
                    }                    
                }
            }

            VA.Documents.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit();
        }

        [TestMethod]
        public void VDX_MultiPageDocument()
        {
            string output_filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");

            var template = new VA.VDX.Template(); // the default template
            var doc = new VA.VDX.Elements.Drawing(template);

            VDX_Files.GetPage01_Simple_Fill_Format(doc);
            VDX_Files.GetPage02_Locking(doc);
            VDX_Files.GetPage03_Text_Block(doc);
            VDX_Files.GetPage04_Simple_Text(doc);
            VDX_Files.GetPage05_Formatted_Text(doc);
            VDX_Files.GetPage06_All_FillPatterns(doc);
            VDX_Files.GetPage08_Connector_With_Geometry(doc);
            VDX_Files.GetPage09_Layout(doc);
            VDX_Files.GetPage10_layers(doc);
            VDX_Files.GetPage11_Add_color(doc);
            VDX_Files.GetPage12_AdjustToTextSize(doc);
            VDX_Files.GetPage13_MultipleConnectors(doc);
            VDX_Files.GetPage14_Hyperlinks(doc);

            var w1 = new VA.VDX.Elements.DocumentWindow();
            w1.ShowGrid = false;
            w1.ShowGuides = false;
            w1.ShowConnectionPoints = false;
            w1.ShowPageBreaks = false;
            w1.Page = 0; // point to first pagees
            
            doc.Windows.Add(w1);

            doc.Save(output_filename);
            
            // Verify this file can be loaded
            CheckIfLoadsWithoutErrorLog(output_filename);
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
            var customprops = CustomPropertyHelper.Get(shape);

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
        public void VDX_CustomTemplate()
        {
            string output_filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");

            var template = new VA.VDX.Template(TestVisioAutomationVDX.Properties.Resources.template_router__vdx);
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

            CheckIfLoadsWithoutErrorLog(output_filename);
        }

        [TestMethod]
        public void VDX_CheckNoErrorOnLoad()
        {
            // This test tends to fail with Visio 2013

            string output_filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");
            System.IO.File.WriteAllText(output_filename, TestVisioAutomationVDX.Properties.Resources.template_router__vdx);

            var app = new IVisio.Application();
            
            DeleteXmlErrorLog(app);
            
            if (XmlErrorLogExists(app))
            {
                Assert.Fail("Before TryOpen: Error log exists and we did not expect it");
            }
            var doc = TryOpen(app.Documents, output_filename);

            var visio_version = app.Version;
            var vermajor = int.Parse(visio_version.Split('.')[0]);
            
            // Prior to Visio 2013 the error log file exists only
            // if there is a problem loading the VDX file
            // in Visio 2013 it is always created

            if (vermajor < 15)
            {
                if (XmlErrorLogExists(app))
                {
                    Assert.Fail("After TryOpen: Error log exists and we did not expect it");
                }
                
            }
            else
            {
                if (!XmlErrorLogExists(app))
                {
                    Assert.Fail("After TryOpen: Error log does not exist");
                }

                string logfile = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);
                string logtext = System.IO.File.ReadAllText(logfile);

                if (logtext.Contains("[Warning]"))
                {
                    Assert.Fail("Error log contains [Warning]");
                }

                if (logtext.Contains("[Error]"))
                {
                    Assert.Fail("Error log contains [Error]");
                }

            }

            VA.Documents.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit();
        }

        [TestMethod]
        public void VDX_CheckErrorOnLoadLogFileExists()
        {
            string output_filename = TestVisioAutomation.Common.Globals.Helper.GetTestMethodOutputFilename(".vdx");
            System.IO.File.WriteAllText(output_filename, TestVisioAutomationVDX.Properties.Resources.vdx_with_errors_1_vdx);

            var app = new IVisio.Application();

            DeleteXmlErrorLog(app);

            if (XmlErrorLogExists(app))
            {
                Assert.Fail("Before TryOpen: Error log exists and we did not expect it");
            }

            // this causes the doc to load no matter what the error ))))))
            var doc = TryOpen(app.Documents, output_filename);

            string logfile = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);

            if (!XmlErrorLogExists(app))
            {
                Assert.Fail("Error log does not exist even though we expected it to");
            }

            string logtext = System.IO.File.ReadAllText(logfile);

            if (!logtext.Contains("[Warning]"))
            {
                Assert.Fail("Error log does not contain [Warning]");
            }

            Assert.AreEqual(1, app.Documents.Count);
            VA.Documents.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit(true);
        }


        public static void DeleteXmlErrorLog(IVisio.Application app)
        {
            string logfilename = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);

            if (logfilename == null)
            {
                // nothing to do
                return;
            }

            if (System.IO.File.Exists(logfilename))
            {
                System.IO.File.Delete(logfilename);
            }
        }

        public static bool XmlErrorLogExists(IVisio.Application app)
        {
            string logfilename = VA.Application.ApplicationHelper.GetXMLErrorLogFilename(app);

            if (logfilename == null)
            {
                return false;
            }

            return System.IO.File.Exists(logfilename);
        }
    }
}