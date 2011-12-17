using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace TestVisioAutomationVDX
{
    [TestClass]
    public class VDX_Tests
    {
        public void CheckIfLoadsWithoutErrorLog(string filename)
        {
            var app = new IVisio.Application();

            DeleteXmlErrorLog(app);

            // this causes the doc to load no matter what the error ))))))
            using (var scope = app.CreateAlertResponseScope(VA.UI.AlertResponseCode.No))
            {
                var doc = app.Documents.Open(filename);
            }

            if (XmlErrorLogExists(app))
            {
                Assert.Fail("XML Error Log Error Was Created when opening the VDX file");
            }

            VA.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit();
        }

        [TestMethod]
        public void CreateSampleVDX1()
        {
            string output_filename = TestCommon.VisioTestCommon.Helper.GetTestMethodOutputFilename(".vdx");

            // First load a starter VDX (a.k.a "the template") - we will build a new VDX from this one
            //var template_dom = System.Xml.Linq.XDocument.Load(template_filename);
            var template_dom = System.Xml.Linq.XDocument.Parse(VA.VDX.Elements.Drawing.DefaultTemplateXML);

            // Clean up the template - remove the existing pages
            VA.VDX.VDXWriter.CleanUpTemplate(template_dom);

            // Create a new Drawing based on the template
            var doc = new VA.VDX.Elements.Drawing(template_dom);

            GetPage01_Simple_Fill_Format(doc);
            GetPage02_Locking(doc);
            GetPage03_Text_Block(doc);
            GetPage04_Simple_Text(doc);
            GetPage05_Formatted_Text(doc);
            GetPage06_All_FillPatterns(doc);
            GetPage07_CustomProps(doc);
            GetPage08_Connector_With_Geometry(doc);
            GetPage09_Layout(doc);
            GetPage10_layers(doc);
            GetPage11_Add_color(doc);

            var w1 = new VA.VDX.Elements.DocumentWindow();
            w1.ShowGrid = false;
            w1.ShowGuides = false;
            w1.ShowConnectionPoints = false;
            w1.ShowPageBreaks = false;
            w1.Page = 1;


            doc.Windows.Add(w1);

            // now create a writer object
            var vdx_writer = new VA.VDX.VDXWriter();
            // merge the template with the new in-memory drawing and save it to the output fie
            vdx_writer.CreateVDX(doc, template_dom, output_filename);

            CheckIfLoadsWithoutErrorLog(output_filename);
        }

        private VA.VDX.Elements.Page GetPage01_Simple_Fill_Format(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            // using that ID draw a rounded rectangle at pinpos(2,2) with size (2.5,2)
            var shape2 = new VA.VDX.Elements.Shape(rounded_rect_id, 2, 2, 2.5, 2);
            page.Shapes.Add(shape2);

            // set the fill properties of the second shape
            shape2.Fill = new VA.VDX.Sections.Fill();
            shape2.Fill.ForegroundColor.Result = 0xff0000;
            shape2.Fill.BackgroundColor.Result = 0x55ff00;
            shape2.Fill.ForegroundTransparency.Result = 0.1;
            shape2.Fill.BackgroundTransparency.Result = 0.9;
            shape2.Fill.Pattern.Result = 40;

            shape1.Line = new VA.VDX.Elements.Line();
            shape1.Line.Weight.Result = 1.0;

            shape1.XForm.Angle.Result = System.Math.PI/4;

            return page;
        }

        private VA.VDX.Elements.Page GetPage02_Locking(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);
            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            // find the id of the master for dynamic connector
            int dynamic_connector_id = doc.GetMasterMetaData("Dynamic Connector").ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            shape1.Text.Add("This shape is completely locked");

            shape1.Protection = new VA.VDX.Sections.Protection();
            shape1.Protection.SetAll(true);
            return page;
        }

        private VA.VDX.Elements.Page GetPage03_Text_Block(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            // find the id of the master for dynamic connector
            int dynamic_connector_id = doc.GetMasterMetaData("Dynamic Connector").ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            shape1.Text.Add("This shape has its text block set");

            shape1.TextBlock = new VA.VDX.Sections.TextBlock();
            shape1.TextBlock.LeftMargin.Result = 0.25;
            shape1.TextBlock.RightMargin.Result = 0.20;
            shape1.TextBlock.TopMargin.Result = 0.1;
            shape1.TextBlock.BottomMargin.Result = 0.15;

            shape1.TextBlock.TextBkgnd.Result = 0xff8800;
            shape1.TextBlock.TextBkgndTrans.Result = 0.5;
            return page;
        }

        private VA.VDX.Elements.Page GetPage04_Simple_Text(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Page4Shape1");

            var shape2 = new VA.VDX.Elements.Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Page4Shape2");

            var shape3 = VA.VDX.Elements.Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;

            shape3.Line = new VA.VDX.Elements.Line();
            shape3.Line.EndArrow.Result = 3;

            page.Shapes.Add(shape3);

            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }

        private VA.VDX.Elements.Page GetPage05_Formatted_Text(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            // using that ID draw a rounded rectangle at pinpos(2,2) with size (2.5,2)
            var shape2 = new VA.VDX.Elements.Shape(rounded_rect_id, 2, 2, 2.5, 2);
            page.Shapes.Add(shape2);

            shape1.XForm.Angle.Result = System.Math.PI/4;

            shape1.Text.Add("HELLO");
            shape2.TextXForm = new VA.VDX.Sections.TextXForm();
            shape2.TextXForm.Width.Formula = "GUARD(TEXTWIDTH(TheText))";
            shape2.TextXForm.Height.Formula = "GUARD(TEXTHEIGHT(TheText,TxtWidth))";
            shape2.TextXForm.PinY.Formula = "-TxtHeight*0.5";

            var font_segoeui = doc.AddFace("Segoe UI");
            var font_gillsans = doc.AddFace("Gill Sans MT");
            var font_trebuchet = doc.AddFace("Trebuchet MS");

            var charfmt1 = new VA.VDX.Sections.Char();
            charfmt1.Font.Result = font_gillsans.ID;
            charfmt1.DoubleUnderline.Result = true;
            charfmt1.Size.Result = 18.0;
            charfmt1.Transparency.Result = 0.5;
            charfmt1.Style.Result = VA.VDX.Enums.CharStyle.Italic | VA.VDX.Enums.CharStyle.Bold |
                                    VA.VDX.Enums.CharStyle.Underline;

            var charfmt2 = new VA.VDX.Sections.Char();
            charfmt2.Font.Result = font_trebuchet.ID;
            charfmt2.Strikethru.Result = true;
            charfmt2.Size.Result = 26;

            var charfmt3 = new VA.VDX.Sections.Char();
            charfmt3.Font.Result = font_segoeui.ID;
            charfmt3.Strikethru.Result = true;
            charfmt3.RTLText.Result = true;

            var parafmt1 = new VA.VDX.Sections.ParagraphFormat();
            parafmt1.HorzAlign.Result = VA.VDX.Enums.ParaHorizontalAlignment.Center;

            var parafmt2 = new VA.VDX.Sections.ParagraphFormat();
            parafmt2.HorzAlign.Result = VA.VDX.Enums.ParaHorizontalAlignment.Right;

            var parafmt3 = new VA.VDX.Sections.ParagraphFormat();
            parafmt3.HorzAlign.Result = VA.VDX.Enums.ParaHorizontalAlignment.Left;

            shape2.CharFormats = new List<VA.VDX.Sections.Char>();
            shape2.ParaFormats = new List<VA.VDX.Sections.ParagraphFormat>();

            shape2.CharFormats.Add(charfmt1);
            shape2.CharFormats.Add(charfmt2);
            shape2.CharFormats.Add(charfmt3);

            shape2.ParaFormats.Add(parafmt1);
            shape2.ParaFormats.Add(parafmt2);
            shape2.ParaFormats.Add(parafmt3);

            shape2.Text.Add("world1\n", 0, 0, null);
            shape2.Text.Add("world2\n", 1, 1, null);
            shape2.Text.Add("world3", 2, 2, null);
            return page;
        }

        private VA.VDX.Elements.Page GetPage06_All_FillPatterns(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 5);
            doc.Pages.Add(page);

            int rect_id = doc.GetMasterMetaData("REctAngle").ID;

            double width = 1.0;
            double height = 1.0;
            int pattern = 0;
            foreach (int row in Enumerable.Range(0, 5))
            {
                foreach (int col in Enumerable.Range(0, 8))
                {
                    double x0 = col*width;
                    double y0 = row*height;

                    double pinx = x0 + width/2.0;
                    double piny = y0 + height/2.0;

                    var shape = new VA.VDX.Elements.Shape(rect_id, pinx, piny, width, height);
                    page.Shapes.Add(shape);
                    shape.Fill = new VA.VDX.Sections.Fill();
                    shape.Fill.ForegroundColor.Result = 0xff0000;
                    shape.Fill.BackgroundColor.Result = 0x55ff00;
                    shape.Fill.Pattern.Result = pattern++;
                    shape.Text.Add(pattern.ToString());
                }
            }

            return page;
        }

        private VA.VDX.Elements.Page GetPage07_CustomProps(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 5);
            doc.Pages.Add(page);

            int rect_id = doc.GetMasterMetaData("REctAngle").ID;

            var shape = new VA.VDX.Elements.Shape(rect_id, 4, 2, 3, 2);
            shape.CustomProps = new List<VA.VDX.Elements.CustomProp>();

            var cp1 = new VA.VDX.Elements.CustomProp(1, "NAMER");

            shape.CustomProps.Add(cp1);

            page.Shapes.Add(shape);

            return page;
        }

        private VA.VDX.Elements.Page GetPage08_Connector_With_Geometry(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("XXX1");

            var shape2 = new VA.VDX.Elements.Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("XXX2");

            var shape3 = VA.VDX.Elements.Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            page.Shapes.Add(shape3);
            shape3.Geom = new VA.VDX.Sections.Geom();
            shape3.Geom.Rows.Add(new VA.VDX.Sections.MoveTo(1, 3));
            shape3.Geom.Rows.Add(new VA.VDX.Sections.LineTo(5, 3));

            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }

        private VA.VDX.Elements.Page GetPage09_Layout(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            var layout = new VA.VDX.Sections.Layout();
            layout.ShapeRouteStyle.Result = VA.VDX.Enums.RouteStyle.TreeEW;

            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("XXX1");

            shape1.Layout = layout;

            var shape2 = new VA.VDX.Elements.Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("XXX2");

            shape2.Layout = shape1.Layout;

            var shape3 = VA.VDX.Elements.Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            page.Shapes.Add(shape3);
            shape3.Geom = new VA.VDX.Sections.Geom();
            shape3.Geom.Rows.Add(new VA.VDX.Sections.MoveTo(1, 3));
            shape3.Geom.Rows.Add(new VA.VDX.Sections.LineTo(5, 3));

            shape3.Layout = shape1.Layout;
            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }

        private VA.VDX.Elements.Page GetPage10_layers(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            var layer0 = page.AddLayer("Layer0", 0);
            var layer1 = page.AddLayer("Layer1", 1);
            var layer2 = page.AddLayer("Layer2", 2);

            string rounded_rect = "Rounded REctAngle";
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            var layout = new VA.VDX.Sections.Layout();
            layout.ShapeRouteStyle.Result = VA.VDX.Enums.RouteStyle.TreeEW;

            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Shape1");

            shape1.Layout = layout;

            var shape2 = new VA.VDX.Elements.Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Shape2");

            shape2.Layout = shape1.Layout;

            var shape3 = VA.VDX.Elements.Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            page.Shapes.Add(shape3);
            shape3.Geom = new VA.VDX.Sections.Geom();
            shape3.Geom.Rows.Add(new VA.VDX.Sections.MoveTo(1, 3));
            shape3.Geom.Rows.Add(new VA.VDX.Sections.LineTo(5, 3));

            shape3.Layout = shape1.Layout;
            page.ConnectShapesViaConnector(shape3, shape1, shape2);

            shape3.LayerMembership = new List<int> {layer0.Index, layer2.Index};
            shape1.LayerMembership = new List<int> {layer1.Index};
            shape2.LayerMembership = new List<int> {layer2.Index};

            return page;
        }

        private VA.VDX.Elements.Page GetPage11_Add_color(VA.VDX.Elements.Drawing doc)
        {
            var page = new VA.VDX.Elements.Page(8, 4);
            doc.Pages.Add(page);

            var layer0 = page.AddLayer("Foo", 0);

            var layer1 = page.AddLayer("BAR", 1);

            string rounded_rect = "Rounded REctAngle";
            int rounded_rect_id = doc.GetMasterMetaData(rounded_rect).ID;

            var layout = new VA.VDX.Sections.Layout();
            layout.ShapeRouteStyle.Result = VA.VDX.Enums.RouteStyle.TreeEW;

            var shape1 = new VA.VDX.Elements.Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);

            doc.Colors.Add(new VA.VDX.Elements.ColorEntry {RGB = 0x123456});
            return page;
        }

        [TestMethod]
        public void CreateSampleVDX2()
        {
            string output_filename = TestCommon.VisioTestCommon.Helper.GetTestMethodOutputFilename(".vdx");

            const string shapeMasterName = "Router";

            var templateDom =
                System.Xml.Linq.XDocument.Parse(TestVisioAutomationVDX.Properties.Resources.template_router__vdx);
            VA.VDX.VDXWriter.CleanUpTemplate(templateDom);
            var doc = new VisioAutomation.VDX.Elements.Drawing(templateDom);
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
            int shapeMasterNameId = doc.GetMasterMetaData(shapeMasterName).ID;
            bool shapeMasterNameGroup = doc.GetMasterMetaData(shapeMasterName).IsGroup;

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

            // add shape3
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

            shape3.LayerMembership = new List<int> {layer0.Index, layer2.Index};
            shape1.LayerMembership = new List<int> {layer1.Index};
            shape2.LayerMembership = new List<int> {layer2.Index};

            // write document to disk as .vdx file
            var vdxWriter = new VA.VDX.VDXWriter();
            vdxWriter.CreateVDX(doc, templateDom, output_filename);

            CheckIfLoadsWithoutErrorLog(output_filename);
        }

        [TestMethod]
        public void CheckNoErrorOnLoad1()
        {
            string output_filename = TestCommon.VisioTestCommon.Helper.GetTestMethodOutputFilename(".vdx");
            System.IO.File.WriteAllText(output_filename, TestVisioAutomationVDX.Properties.Resources.template_router__vdx);

            var app = new IVisio.Application();
            
            DeleteXmlErrorLog(app);

            // this causes the doc to load no matter what the error ))))))
            using (var scope = app.CreateAlertResponseScope(VA.UI.AlertResponseCode.No))
            {
                var doc = app.Documents.Open(output_filename);
            }

            if (XmlErrorLogExists(app))
            {
                Assert.Fail("Error log exists and we did not expect it");
            }
            VA.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit();
        }

        [TestMethod]
        public void CheckErrorOnLoad1()
        {
            string output_filename = TestCommon.VisioTestCommon.Helper.GetTestMethodOutputFilename(".vdx");
            System.IO.File.WriteAllText(output_filename, TestVisioAutomationVDX.Properties.Resources.vdx_with_errors_1_vdx);

            var app = new IVisio.Application();

            DeleteXmlErrorLog(app);

            // this causes the doc to load no matter what the error ))))))
            using (var scope = app.CreateAlertResponseScope(VA.UI.AlertResponseCode.No))
            {
                var doc = app.Documents.Open(output_filename);
            }

            if (!XmlErrorLogExists(app))
            {
                System.Windows.Forms.MessageBox.Show("Error log does not exist even though we expected it to");
            }
            Assert.AreEqual(1, app.Documents.Count);
            VA.DocumentHelper.ForceCloseAll(app.Documents);
            app.Quit();
        }

        public static void DeleteXmlErrorLog(IVisio.Application app)
        {
            string logfilename = VA.ApplicationHelper.GetXMLErrorLogFilename(app);

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
            string logfilename = VA.ApplicationHelper.GetXMLErrorLogFilename(app);

            if (logfilename == null)
            {
                return false;
            }

            return System.IO.File.Exists(logfilename);
        }
    }
}