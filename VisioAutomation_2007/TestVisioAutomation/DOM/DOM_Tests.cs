using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class DOM_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Text_Markup_CharacterFormatting()
        {
            this.MarkupCharacterBold();
            this.MarkupCharacterComplex();
            this.MarkupCharacterFont();
            this.MarkupCharacterItalic();
            this.MarkupCharacterPlain();
            this.MarkupParagraphCenter();
            this.MarkupParagraphDefault();
            this.MarkupParagraphLeft();
            this.MarkupParagraphRight();
        }

        public void MarkupCharacterPlain()
        {
            var m = new VA.Text.Markup.TextElement("{Normal}");
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);

            page1.Delete(0);
        }

        public void MarkupCharacterBold()
        {
            var m = new VA.Text.Markup.TextElement("{Bold}");
            m.CharacterCells.Style = (int)VA.Text.CharStyle.Bold;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, charfmt[0].Style.Result);

            page1.Delete(0);
        }

        public void MarkupCharacterItalic()
        {
            var m = new VA.Text.Markup.TextElement("{Italic}");
            m.CharacterCells.Style = (int)VA.Text.CharStyle.Italic;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, charfmt[0].Style.Result);

            page1.Delete(0);
        }

        public void MarkupCharacterFont()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));

            var impact = page1.Document.Fonts["Arial"];
            var m = new VA.Text.Markup.TextElement("Normal Text in Impact Font");
            m.CharacterCells.Font = impact.ID;
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);
            Assert.AreEqual(0, charfmt[0].Style.Result);
            Assert.AreEqual(impact.ID, charfmt[0].Font.Result);

            page1.Delete(0);
        }

        public void MarkupCharacterComplex()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var doc = page1.Document;
            var fonts = doc.Fonts;

            var segoeui = fonts["Segoe UI"];
            var impact = fonts["Arial"];
            var couriernew = fonts["Courier New"];
            var georgia = fonts["Georgia"];

            var t1 = new VA.Text.Markup.TextElement("{Normal}");
            t1.CharacterCells.Font = segoeui.ID;

            var t2 = t1.AddElement("{Italic}");
            t2.CharacterCells.Style = (int)VA.Text.CharStyle.Italic;
            t2.CharacterCells.Font = impact.ID;

            var t3 = t2.AddElement("{Bold}");
            t3.CharacterCells.Style = (int)VA.Text.CharStyle.Bold;
            t3.CharacterCells.Font = couriernew.ID;

            var t4 = t2.AddElement("{Bold Italic}");
            t4.CharacterCells.Style = (int)(VA.Text.CharStyle.Bold | VA.Text.CharStyle.Italic);
            t4.CharacterCells.Font = georgia.ID;

            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            t1.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;

            // check the number of character regions
            Assert.AreEqual(5, charfmt.Count);

            // check the fonts
            Assert.AreEqual(segoeui.ID, charfmt[0].Font.Result);
            Assert.AreEqual(impact.ID, charfmt[1].Font.Result);
            Assert.AreEqual(couriernew.ID, charfmt[2].Font.Result);
            Assert.AreEqual(georgia.ID, charfmt[3].Font.Result);
            Assert.AreEqual(segoeui.ID, charfmt[4].Font.Result);


            // check the styles
            Assert.AreEqual((int)VA.Text.CharStyle.None, charfmt[0].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, charfmt[1].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, charfmt[2].Style.Result);
            Assert.AreEqual((int)(VA.Text.CharStyle.Italic | VA.Text.CharStyle.Bold), charfmt[3].Style.Result);
            Assert.AreEqual((int)(VA.Text.CharStyle.None), charfmt[4].Style.Result);

            // check the text run content
            var charruns = textfmt.CharacterTextRuns;
            Assert.AreEqual(4, charruns.Count);
            Assert.AreEqual("{Normal}", charruns[0].Text);
            Assert.AreEqual("{Italic}", charruns[1].Text);
            Assert.AreEqual("{Bold}", charruns[2].Text);
            Assert.AreEqual("{Bold Italic}", charruns[3].Text);

            // cleanup
            page1.Delete(0);
        }

        public void MarkupParagraphDefault()
        {
            var m = new VA.Text.Markup.TextElement("{DefaultPara}");
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            page1.Delete(0);
        }

        public void MarkupParagraphLeft()
        {
            var m = new VA.Text.Markup.TextElement("{LeftHAlign}");
            m.ParagraphCells.HorizontalAlign = (int)VA.Drawing.AlignmentHorizontal.Left;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            Assert.AreEqual((int)VA.Drawing.AlignmentHorizontal.Left, parafmt[0].HorizontalAlign.Result);

            page1.Delete(0);
        }

        public void MarkupParagraphCenter()
        {
            var m = new VA.Text.Markup.TextElement("{CenterHAlign}");
            m.ParagraphCells.HorizontalAlign = (int)VA.Drawing.AlignmentHorizontal.Center;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            Assert.AreEqual((int)VA.Drawing.AlignmentHorizontal.Center, parafmt[0].HorizontalAlign.Result);

            page1.Delete(0);
        }

        public void MarkupParagraphRight()
        {
            var m = new VA.Text.Markup.TextElement("{RightHAlign}");
            m.ParagraphCells.HorizontalAlign = (int)VA.Drawing.AlignmentHorizontal.Right;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            Assert.AreEqual((int)VA.Drawing.AlignmentHorizontal.Right, parafmt[0].HorizontalAlign.Result);

            page1.Delete(0);
        }

        public int get_doc_count(IVisio.Application app)
        {
            // get the number of actual drawings, not including templates, stencils, etc.
            var documents = app.Documents;
            var drawings = documents.AsEnumerable()
                .Where(doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing);
            return drawings.Count();
        }

        [TestMethod]
        public void DOM_EmptyRendering()
        {
            // Rendering a DOM should not change the page count
            // Empty DOMs do not add any shapes
            var app = this.GetVisioApplication();

            var page_node = new VA.DOM.Page();
            var doc = this.GetNewDoc();
            page_node.Render(app.ActiveDocument);
            Assert.AreEqual(0, app.ActivePage.Shapes.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_RenderPageToDocument()
        {
            // Rendering a dom page to a document should create a new page
            var app = this.GetVisioApplication();
            var page_node = new VA.DOM.Page();
            var visdoc = this.GetNewDoc();
            Assert.AreEqual(1, visdoc.Pages.Count);
            var page = page_node.Render(app.ActiveDocument);
            Assert.AreEqual(2, visdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_RenderDocumentToApplication()
        {
            // Rendering a dom document to an appliction instance should create a new document
            var app = this.GetVisioApplication();
            var doc_node = new VA.DOM.Document();
            var page_node = new VA.DOM.Page();
            doc_node.Pages.Add(page_node);
            int old_count = app.Documents.Count;
            var newdoc = doc_node.Render(app);
            Assert.AreEqual(old_count + 1, app.Documents.Count);
            Assert.AreEqual(1, newdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_DrawSimpleShape()
        {
            // Create the doc
            var page_node = new VA.DOM.Page();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");
            vrect1.Cells.FillForegnd = "rgb(255,0,0)";
            page_node.Shapes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            VisioAutomationTest.SetPageSize(app.ActivePage, new VA.Drawing.Size(10, 10));
            var page = page_node.Render(app.ActiveDocument);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_DropShapes()
        {
            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var stencil = app.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];


            // Create the doc
            var shape_nodes = new VA.DOM.ShapeList();

            shape_nodes.DrawRectangle(0, 0, 1, 1);
            shape_nodes.Drop(rectmaster, 3, 3);

            shape_nodes.Render(app.ActivePage);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_CustomProperties()
        {
            // Create the doc
            var shape_nodes = new VA.DOM.ShapeList();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");

            vrect1.CustomProperties = new Dictionary<string, VA.Shapes.CustomProperties.CustomPropertyCells>();

            var cp1 = new VA.Shapes.CustomProperties.CustomPropertyCells();
            cp1.Value = "FOOVALUE";
            cp1.Label = "Foo Label";

            var cp2 = new VA.Shapes.CustomProperties.CustomPropertyCells();
            cp2.Value = "BARVALUE";
            cp2.Label = "Bar Label";

            vrect1.CustomProperties["FOO"] = cp1;
            vrect1.CustomProperties["BAR"] = cp2;

            shape_nodes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            shape_nodes.Render(app.ActivePage);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);
            Assert.IsTrue(VA.Shapes.CustomProperties.CustomPropertyHelper.Contains(vrect1.VisioShape, "FOO"));
            Assert.IsTrue(VA.Shapes.CustomProperties.CustomPropertyHelper.Contains(vrect1.VisioShape, "BAR"));

            doc.Close(true);
        }

        [TestMethod]
        public void DOM_DrawOrgChart()
        {
            // How to draw using a Template instead of a doc and a stencil
            const string orgchart_vst = "orgch_u.vst" +
                                        "";

            var app = this.GetVisioApplication();
            var doc_node = new VA.DOM.Document(orgchart_vst, IVisio.VisMeasurementSystem.visMSUS);
            var page_node = new VA.DOM.Page();
            doc_node.Pages.Add(page_node);

            // Have to be smart about selecting the right master with Visio 2013
            int vis_ver = int.Parse(app.Version.Split('.')[0]);
            string position_master_name = vis_ver >= 15 ? "Position Belt" : "Position";

            var s1 = new VisioAutomation.DOM.Shape(position_master_name, null, new VA.Drawing.Point(3, 4));
            page_node.Shapes.Add(s1);
            var doc = doc_node.Render(app);

            doc.Close(true);
        }

        [TestMethod]
        public void DOM_DrawEmpty()
        {
            // Verify that an empty DOM page can be created and rendered
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            page_node.Size = new VA.Drawing.Size(5, 5);
            var page = page_node.Render(doc);

            Assert.AreEqual(0, page.Shapes.Count);
            Assert.AreEqual(new VA.Drawing.Size(5, 5), VisioAutomationTest.GetPageSize(page));

            page.Delete(0);
            doc.Close(true);
        }

        public void DOM_DrawLine()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var line_node_0 = page_node.Shapes.DrawLine(1, 1, 3, 3);
            var page = page_node.Render(doc);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, line_node_0.VisioShapeID);
            Assert.IsNotNull(line_node_0.VisioShape);
            Assert.AreEqual(2.0, line_node_0.VisioShape.CellsU["PinX"].Result[IVisio.VisUnitCodes.visNumber]);
            page.Delete(0);
            doc.Close(true);
        }

        public void DOM_DrawBezier()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var bez_node_0 = page_node.Shapes.DrawBezier(new double[] { 1, 2, 3, 3, 6, 3, 3, 4 });

            var page = page_node.Render(doc);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, bez_node_0.VisioShapeID);
            Assert.IsNotNull(bez_node_0.VisioShape);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_DropMaster()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", 5, 5);

            var page = page_node.Render(doc);

            Assert.AreEqual(2, page.Shapes.Count);

            // Verify that the shapes created both have IDs and shape objects associated with them
            Assert.AreNotEqual(0, master_node_0.VisioShapeID);
            Assert.AreNotEqual(0, master_node_1.VisioShapeID);
            Assert.IsNotNull(master_node_0.VisioShape);
            Assert.IsNotNull(master_node_1.VisioShape);
            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_FormatShape()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var bez_node_0 = page_node.Shapes.DrawBezier(new double[] { 1, 2, 3, 3, 6, 3, 3, 4 });
            var line_node_0 = page_node.Shapes.DrawLine(1, 1, 3, 3);

            master_node_0.Cells.LineWeight = 0.1;
            bez_node_0.Cells.LineWeight = 0.3;
            line_node_0.Cells.LineWeight = 0.5;

            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);
            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_ConnectShapes()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var basic_stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var basic_masters = basic_stencil.Masters;

            var connectors_stencil = doc.Application.Documents.OpenStencil("connec_u.vss");
            var connectors_masters = connectors_stencil.Masters;

            var master1 = basic_masters["Rectangle"];
            var master2 = connectors_masters["Dynamic Connector"];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(master1, 6, 5);
            var dc = page_node.Shapes.Connect(master2, master_node_0, master_node_1);

            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_ConnectShapes2()
        {
            // Deferred means that the stencils (and thus masters) are loaded when rendering
            // and are no loaded by the caller before Render() is called

            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", 3, 3);
            var master_node_1 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", 6, 5);
            var dc = page_node.Shapes.Connect("Dynamic Connector", "connec_u.vss", master_node_0, master_node_1);
            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_VerifyThatUnknownMastersAreDetected()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop("RectangleXXX", "basic_u.vss", 3, 3);

            IVisio.Page page=null;
            bool caught = false;
            try
            {
                page = page_node.Render(doc);
            }
            catch (VA.AutomationException)
            {
                caught = true;
            }
            
            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }
            
            if (page!=null)
            {
                page.Delete(0);
            }
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_VerifyThatUnknownStencilsAreDetected()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop("Rectangle", "basic_uXXX.vss", 3, 3);

            IVisio.Page page = null;
            bool caught = false;
            try
            {
                page = page_node.Render(doc);
            }
            catch (VA.AutomationException)
            {
                caught = true;
            }
            
            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }

            if (page!=null)
            {
                page.Delete(0);                
            }
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_DrawAndDrop()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();

            var rect0 = new VA.Drawing.Rectangle(3, 4, 7, 8);
            var rect1 = new VA.Drawing.Rectangle(8, 1, 9, 5);
            var dropped_shape0 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", rect0);
            var drawn_shape0 = page_node.Shapes.DrawRectangle(rect0);

            var dropped_shape1 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", rect1);
            var drawn_shape1 = page_node.Shapes.DrawRectangle(rect1);

            var page = page_node.Render(doc);

            var xfrms = VA.Shapes.XFormCells.GetCells(page,
                                                        new int[] { dropped_shape0.VisioShapeID, 
                                                            drawn_shape0.VisioShapeID, 
                                                            dropped_shape1.VisioShapeID, 
                                                            drawn_shape1.VisioShapeID });

            Assert.AreEqual(xfrms[1].PinX, xfrms[0].PinX);
            Assert.AreEqual(xfrms[1].PinY, xfrms[0].PinY);

            Assert.AreEqual(xfrms[1].Width, xfrms[0].Width);
            Assert.AreEqual(xfrms[1].Height, xfrms[0].Height);

            Assert.AreEqual(xfrms[3].PinX,   xfrms[2].PinX);
            Assert.AreEqual(xfrms[3].PinY,   xfrms[2].PinY);
            Assert.AreEqual(xfrms[3].Width,  xfrms[2].Width);
            Assert.AreEqual(xfrms[3].Height, xfrms[2].Height);

            doc.Close(true);
        }
    }
}