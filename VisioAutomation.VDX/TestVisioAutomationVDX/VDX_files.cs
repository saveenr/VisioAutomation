using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.VDX.Elements;
using VisioAutomation.VDX.Enums;
using VisioAutomation.VDX.Sections;
using VA = VisioAutomation;

namespace TestVisioAutomationVDX
{
    public class VDX_Files
    {

        public static Page GetPage01_Simple_Fill_Format(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);
           
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            // using that ID draw a rounded rectangle at pinpos(2,2) with size (2.5,2)
            var shape2 = new Shape(rounded_rect_id, 2, 2, 2.5, 2);
            page.Shapes.Add(shape2);

            // set the fill properties of the second shape
            shape2.Fill = new Fill();
            shape2.Fill.ForegroundColor.Result = 0xff0000;
            shape2.Fill.BackgroundColor.Result = 0x55ff00;
            shape2.Fill.ForegroundTransparency.Result = 0.1;
            shape2.Fill.BackgroundTransparency.Result = 0.9;
            shape2.Fill.Pattern.Result = 40;

            shape1.Line = new Line();
            shape1.Line.Weight.Result = 1.0;

            shape1.XForm.Angle.Result = Math.PI/4;

            return page;
        }

        public static Page GetPage02_Locking(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);
            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            // find the id of the master for dynamic connector
            int dynamic_connector_id = doc.GetMasterMetaData("Dynamic Connector").ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            shape1.Text.Add("This shape is completely locked");

            shape1.Protection = new Protection();
            shape1.Protection.SetAll(true);
            return page;
        }

        public static Page GetPage03_Text_Block(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the masters
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            shape1.Text.Add("This shape has its text block set");

            shape1.TextBlock = new TextBlock();
            shape1.TextBlock.LeftMargin.Result = 0.25;
            shape1.TextBlock.RightMargin.Result = 0.20;
            shape1.TextBlock.TopMargin.Result = 0.1;
            shape1.TextBlock.BottomMargin.Result = 0.15;

            shape1.TextBlock.TextBkgnd.Result = 0xff8800;
            shape1.TextBlock.TextBkgndTrans.Result = 0.5;
            return page;
        }

        public static Page GetPage04_Simple_Text(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Page4Shape1");

            var shape2 = new Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Page4Shape2");

            var shape3 = Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;

            shape3.Line = new Line();
            shape3.Line.EndArrow.Result = 3;

            page.Shapes.Add(shape3);

            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }

        public static Page GetPage05_Formatted_Text(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            // using that ID draw a rounded rectangle at pinpos(4,3)
            var shape1 = new Shape(rounded_rect_id, 4, 3);
            page.Shapes.Add(shape1);

            // using that ID draw a rounded rectangle at pinpos(2,2) with size (2.5,2)
            var shape2 = new Shape(rounded_rect_id, 2, 2, 2.5, 2);
            page.Shapes.Add(shape2);

            shape1.XForm.Angle.Result = Math.PI/4;

            shape1.Text.Add("HELLO");
            shape2.TextXForm = new TextXForm();
            shape2.TextXForm.PinY.Formula = "-TxtHeight*0.5";

            var font_segoeui = doc.AddFace("Segoe UI");
            var font_gillsans = doc.AddFace("Gill Sans MT");
            var font_trebuchet = doc.AddFace("Trebuchet MS");

            var charfmt1 = new VisioAutomation.VDX.Sections.Char();
            charfmt1.Font.Result = font_gillsans.ID;
            charfmt1.DoubleUnderline.Result = true;
            charfmt1.Size.Result = 18.0;
            charfmt1.Transparency.Result = 0.5;
            charfmt1.Style.Result = CharStyle.Italic | CharStyle.Bold |
                                    CharStyle.Underline;

            var charfmt2 = new VisioAutomation.VDX.Sections.Char();
            charfmt2.Font.Result = font_trebuchet.ID;
            charfmt2.Strikethru.Result = true;
            charfmt2.Size.Result = 26;

            var charfmt3 = new VisioAutomation.VDX.Sections.Char();
            charfmt3.Font.Result = font_segoeui.ID;
            charfmt3.Strikethru.Result = true;
            charfmt3.RTLText.Result = true;

            var parafmt1 = new ParagraphFormat();
            parafmt1.HorzAlign.Result = ParaHorizontalAlignment.Center;

            var parafmt2 = new ParagraphFormat();
            parafmt2.HorzAlign.Result = ParaHorizontalAlignment.Right;

            var parafmt3 = new ParagraphFormat();
            parafmt3.HorzAlign.Result = ParaHorizontalAlignment.Left;

            shape2.CharFormats = new List<VisioAutomation.VDX.Sections.Char>();
            shape2.ParaFormats = new List<ParagraphFormat>();

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

        public static Page GetPage06_All_FillPatterns(Drawing doc)
        {
            var page = new Page(8, 5);
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

                    var shape = new Shape(rect_id, pinx, piny, width, height);
                    page.Shapes.Add(shape);
                    shape.Fill = new Fill();
                    shape.Fill.ForegroundColor.Result = 0xff0000;
                    shape.Fill.BackgroundColor.Result = 0x55ff00;
                    shape.Fill.Pattern.Result = pattern++;
                    shape.Text.Add(pattern.ToString());
                }
            }

            return page;
        }

        public static Page GetPage08_Connector_With_Geometry(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("XXX1");

            var shape2 = new Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("XXX2");

            var shape3 = Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            page.Shapes.Add(shape3);
            shape3.Geom = new Geom();
            shape3.Geom.Rows.Add(new MoveTo(1, 3));
            shape3.Geom.Rows.Add(new LineTo(5, 3));

            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }

        public static Page GetPage09_Layout(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var layout = new Layout();
            layout.ShapeRouteStyle.Result = RouteStyle.TreeEW;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("XXX1");

            shape1.Layout = layout;

            var shape2 = new Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("XXX2");

            shape2.Layout = shape1.Layout;

            var shape3 = Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            page.Shapes.Add(shape3);
            shape3.Geom = new Geom();
            shape3.Geom.Rows.Add(new MoveTo(1, 3));
            shape3.Geom.Rows.Add(new LineTo(5, 3));

            shape3.Layout = shape1.Layout;
            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }

        public static Page GetPage10_layers(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            var layer0 = page.AddLayer("Layer0", 0);
            var layer1 = page.AddLayer("Layer1", 1);
            var layer2 = page.AddLayer("Layer2", 2);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var layout = new Layout();
            layout.ShapeRouteStyle.Result = RouteStyle.TreeEW;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Shape1");

            shape1.Layout = layout;

            var shape2 = new Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Shape2");

            shape2.Layout = shape1.Layout;

            var shape3 = Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            page.Shapes.Add(shape3);
            shape3.Geom = new Geom();
            shape3.Geom.Rows.Add(new MoveTo(1, 3));
            shape3.Geom.Rows.Add(new LineTo(5, 3));

            shape3.Layout = shape1.Layout;
            page.ConnectShapesViaConnector(shape3, shape1, shape2);

            shape3.LayerMembership = new List<int> {layer0.Index, layer2.Index};
            shape1.LayerMembership = new List<int> {layer1.Index};
            shape2.LayerMembership = new List<int> {layer2.Index};

            return page;
        }

        public static Page GetPage11_Add_color(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            var layer0 = page.AddLayer("Foo", 0);
            var layer1 = page.AddLayer("BAR", 1);

            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var layout = new Layout();
            layout.ShapeRouteStyle.Result = RouteStyle.TreeEW;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);

            doc.Colors.Add(new ColorEntry {RGB = 0x123456});
            return page;
        }

        public static Page GetPage12_AdjustToTextSize(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Page12Shape1");

            shape1.XForm.Width.Formula = "GUARD(TEXTWIDTH(TheText))";
            shape1.XForm.Height.Formula = "GUARD(TEXTHEIGHT(TheText,Width))";

            return page;
        }

        public static Page GetPage13_MultipleConnectors(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            // Add the first shape
            var shape1 = new Shape(rounded_rect_id, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Page13Shape1");

            // Add the second shape
            var shape2 = new Shape(rounded_rect_id, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Page13Shape2");

            // Add the Connector
            var shape3 = Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.EndY.Result = 0;
            shape3.Line = new Line();
            shape3.Line.EndArrow.Result = 3;
            page.Shapes.Add(shape3);

            page.ConnectShapesViaConnector(shape3, shape1, shape2);

            // Add the Connector
            var shape4 = Shape.CreateDynamicConnector(doc);
            shape4.XForm1D.EndY.Result = 0;
            shape4.Line = new Line();
            shape4.Line.EndArrow.Result = 3;
            page.Shapes.Add(shape4);

            page.ConnectShapesViaConnector(shape4, shape1, shape2);

            return page;
        }

        public static Page GetPage14_Hyperlinks(Drawing doc)
        {
            var page = new Page(8, 4);
            doc.Pages.Add(page);

            // find the id of the master for rounded rectangles
            int rounded_rect_id = doc.GetMasterMetaData("Rounded REctAngle").ID;

            var shape1 = new Shape(rounded_rect_id, 1, 3);
            shape1.Text.Add("No Hyperlinks");

            var shape2 = new Shape(rounded_rect_id, 5, 3);
            shape2.Text.Add("1 Hyperlink");
            shape2.Hyperlinks = new List<Hyperlink>();
            shape2.Hyperlinks.Add(new Hyperlink("Google", "http://google.com"));

            var shape3 = new Shape(rounded_rect_id, 5, 3);
            shape3.Text.Add("2 Hyperlinks");
            shape3.Hyperlinks = new List<Hyperlink>();
            shape3.Hyperlinks.Add(new Hyperlink("Google", "http://google.com"));
            shape3.Hyperlinks.Add(new Hyperlink("Microsoft", "http://microsoft.com"));

            page.Shapes.Add(shape1);
            page.Shapes.Add(shape2);
            page.Shapes.Add(shape3);

            page.ConnectShapesViaConnector(shape3, shape1, shape2);
            return page;
        }
    }
}