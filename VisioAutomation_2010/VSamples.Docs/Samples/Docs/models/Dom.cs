using IVisio = Microsoft.Office.Interop.Visio;
using VADOM = VisioAutomation.Models.Dom;
using VATEXT = VisioAutomation.Models.Text;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/dom.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsDomSamples
    {
        public static void Sample_01(IVisio.Page page, System.Collections.Generic.List<IVisio.Shape> shapes, IVisio.Page visioPage, IVisio.Document visioDoc, IVisio.Application visioApp, double a, VADOM.Page page_node, VADOM.ShapeList shape_list)
        {
            // <sample id="models/dom#01" doc="models/dom.md" line="38">
            // New document (blank here; the constructor can also take a template)
            var doc_node = new VADOM.Document();
            doc_node.Pages.Add(page_node);
            IVisio.Document newDoc = doc_node.Render(visioApp);

            // New page in a document you already have
            IVisio.Page newPage = page_node.Render(visioDoc);

            // An existing page, including its page-level settings
            page_node.Render(visioPage);

            // Only the shapes, onto an existing page
            shape_list.Render(visioPage);
            // </sample>
        }

        public static void Sample_02(IVisio.Document visioDoc)
        {
            // <sample id="models/dom#02" doc="models/dom.md" line="67">
            var page_node = new VADOM.Page();
            var rect = new VADOM.Rectangle(1, 1, 9, 9);
            rect.Text = new VATEXT.Element("Hello, Visio");
            rect.Cells.FillForeground = "rgb(255,0,0)";
            page_node.Shapes.Add(rect);

            page_node.Render(visioDoc);   // visioDoc is an IVisio.Document
            // </sample>
        }

        public static void Sample_03(IVisio.Master master, IVisio.Document visioDoc)
        {
            // <sample id="models/dom#03" doc="models/dom.md" line="86">
            var page_node = new VADOM.Page();

            // By master object (master already resolved)
            var stencil = visioDoc.Application.Documents.OpenStencil("basic_u.vss");
            var rectMaster = stencil.Masters["Rectangle"];
            page_node.Shapes.Drop(rectMaster, 3, 3);

            // By name (resolved at render time)
            page_node.Shapes.Drop("Rectangle", "basic_u.vss", 5, 5);

            page_node.Render(visioDoc);
            // </sample>
        }

        public static void Sample_04(IVisio.Document visioDoc)
        {
            // <sample id="models/dom#04" doc="models/dom.md" line="108">
            var page_node = new VADOM.Page();
            var rectMaster = "Rectangle";
            var rectStencil = "basic_u.vss";

            var s1 = new VADOM.Shape(rectMaster, rectStencil, new VisioAutomation.Core.Point(2, 2));
            var s2 = new VADOM.Shape(rectMaster, rectStencil, new VisioAutomation.Core.Point(6, 6));
            page_node.Shapes.Add(s1);
            page_node.Shapes.Add(s2);

            page_node.Shapes.Connect("Dynamic Connector", "connec_u.vss", s1, s2);

            page_node.Render(visioDoc);
            // </sample>
        }

        public static void Sample_05()
        {
            // <sample id="models/dom#05" doc="models/dom.md" line="129">
            var rect = new VADOM.Rectangle(1, 1, 4, 4);
            rect.CustomProperties = new VisioAutomation.Shapes.CustomPropertyDictionary();

            var owner = new VisioAutomation.Shapes.CustomPropertyCells();
            owner.SetString("Alex");          // typed setter handles formula encoding
            owner.Label = "\"Owner\"";        // pre-encoded literal

            rect.CustomProperties["Owner"] = owner;
            // </sample>
        }

        public static void Sample_06(IVisio.Document visioDoc, VADOM.Page page_node)
        {
            // <sample id="models/dom#06" doc="models/dom.md" line="146">
            page_node.Render(visioDoc);

            foreach (var s in page_node.Shapes)
            {
                if (s is VADOM.Rectangle r && r.VisioShape != null)
                {
                    // Drop into the imperative API for anything DOM doesn't cover.
                    r.VisioShape.Text = "post-render touch-up";
                }
            }
            // </sample>
        }
    }
}
