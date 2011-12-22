using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;

namespace VisioAutomationSamples
{
    public static class ConnectorSamples
    {
        public static void ConnectorsToBack()
        {
            var doc = SampleEnvironment.Application.ActiveDocument;
            var pages = doc.Pages;
            var page = pages.Add();

            // get the data and the labels to use
            var data = new double[] {1, 2, 3, 4, 5, 6};

            var radius = 3.0;
            var center = new VA.Drawing.Point(4, 4);
            var slices = VA.Layout.Radial.PieSlice.GetSlicesFromValues(center, radius, data);
            foreach (var slice in slices)
            {
                slice.Render(page);
            }

            // based on this example: http://www.visguy.com/2009/06/17/send-all-connectors-to-back/

            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Dynamic Connector"];
            var connector = page.Drop(master, 0, 0);
            var r1 = page.DrawRectangle(0, 1, 2, 2);
            var r2 = page.DrawRectangle(7, 7, 8, 8);

            VA.Connections.ConnectorHelper.ConnectShapes(connector, r1, r2);

            var con_layer = page.Layers["Connector"];
            var sel = VA.SelectionHelper.SelectShapesInLayer(page, con_layer);
            sel.SendToBack();
        }
    }
}