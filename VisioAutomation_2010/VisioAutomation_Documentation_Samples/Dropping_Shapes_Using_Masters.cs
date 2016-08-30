using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Documentation_Samples
{
    public static class Dropping_Shapes_Using_Masters
    {
        public static void One_shape_at_a_time(IVisio.Document doc)
        {
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];
            var page = doc.Pages.Add();

            var shape1 = page.Drop(rectmaster, 1.0, 2.0);

            var p = new VisioAutomation.Drawing.Point(5.0, 4.0);
            var shape2 = page.Drop(rectmaster, p);

            //cleanup
            page.Delete(0);
        }

        public static void Multiple_shapes_at_a_time(IVisio.Document doc)
        {
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];
            var page = doc.Pages.Add();

            var centerpoints = new[] {
                new VisioAutomation.Drawing.Point(1, 2),
                new VisioAutomation.Drawing.Point(5, 4)
            };
            var masters = new[] { rectmaster, rectmaster };
            short[] shapeids = page.DropManyU(masters, centerpoints);

            //cleanup
            page.Delete(0);
        }

    }
}