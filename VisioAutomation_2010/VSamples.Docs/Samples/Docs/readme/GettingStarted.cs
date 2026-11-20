using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page readme/getting-started.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ReadmeGettingStartedSamples
    {
        public static void Sample_01()
        {
            // <sample id="readme/getting-started#01" doc="readme/getting-started.md" line="39">
            var app = new Microsoft.Office.Interop.Visio.Application();
            // </sample>
        }

        public static void Sample_02()
        {
            // <sample id="readme/getting-started#02" doc="readme/getting-started.md" line="47">
            var app = new Microsoft.Office.Interop.Visio.Application();
            var doc = app.Documents.Add("");
            // </sample>
        }

        public static void Sample_03(double a)
        {
            // <sample id="readme/getting-started#03" doc="readme/getting-started.md" line="54">
            var app = new Microsoft.Office.Interop.Visio.Application();
            var doc = app.Documents.Add("");
            var page = app.ActivePage; // Creating a new doc auto-creates an empty active page.
            var s1 = page.DrawRectangle(0, 0, 5, 3);
            s1.Text = "Hello World";
            // </sample>
        }

        public static void Sample_04()
        {
            // <sample id="readme/getting-started#04" doc="readme/getting-started.md" line="70">

            // </sample>
        }

        public static void Sample_05()
        {
            // <sample id="readme/getting-started#05" doc="readme/getting-started.md" line="82">
            var app = new Microsoft.Office.Interop.Visio.Application();
            var doc = app.Documents.Add("");
            var page = app.ActivePage;
            var stencil = app.Documents.OpenStencil("basic_u.vss");
            var masters = stencil.Masters;
            var rectangle_master = masters["Rectangle"];
            page.Drop(rectangle_master, 3, 3);
            // </sample>
        }

        public static void Sample_06(IVisio.Document doc)
        {
            // <sample id="readme/getting-started#06" doc="readme/getting-started.md" line="102">
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];
            var page2 = doc.Pages.Add();

            var centerpoints = new[] {
                new VisioAutomation.Core.Point(1, 2),
                new VisioAutomation.Core.Point(5, 4),
            };
            var masters = new[] { rectmaster, rectmaster };
            short[] shapeids = page2.DropManyU(masters, centerpoints);
            // </sample>
        }

        public static void Sample_07(IVisio.Document doc)
        {
            // <sample id="readme/getting-started#07" doc="readme/getting-started.md" line="119">
            doc.SaveAs(@"C:\Temp\hello.vsdx");
            // </sample>
        }

        public static void Sample_08(IVisio.Shape shape, System.Collections.Generic.List<IVisio.Shape> shapes, double a)
        {
            // <sample id="readme/getting-started#08" doc="readme/getting-started.md" line="129">
            var app = new IVisio.Application();
            var doc = app.Documents.Add("");
            var page = app.ActivePage;

            // Manual draw
            var s1 = page.DrawRectangle(0, 0, 5, 3);
            s1.Text = "Hello World";

            // Drop a single shape from a stencil
            var stencil = app.Documents.OpenStencil("basic_u.vss");
            var rectangle_master = stencil.Masters["Rectangle"];
            page.Drop(rectangle_master, 7, 3);

            // Drop many shapes in a single batched call
            var page2 = doc.Pages.Add();
            var centerpoints = new[] {
                new VisioAutomation.Core.Point(1, 2),
                new VisioAutomation.Core.Point(5, 4),
            };
            var masters = new[] { rectangle_master, rectangle_master };
            short[] shapeids = page2.DropManyU(masters, centerpoints);

            doc.SaveAs(@"C:\Temp\hello.vsdx");
            // </sample>
        }
    }
}
