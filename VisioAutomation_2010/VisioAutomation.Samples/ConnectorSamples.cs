using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes;

namespace VisioAutomationSamples
{
    public static class ConnectorSamples
    {
        public static void ConnectorsToBack()
        {
            var doc = SampleEnvironment.Application.ActiveDocument;
            var pages = doc.Pages;
            var page = pages.Add();
            
            // based on this example: http://www.visguy.com/2009/06/17/send-all-connectors-to-back/

            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Dynamic Connector"];
            var r0 = page.DrawRectangle(3,3,5,5);
            var r1 = page.DrawRectangle(0, 1, 2, 2);
            var r2 = page.DrawRectangle(7, 7, 8, 8);
            var connector = page.Drop(master, 0, 0);
            ConnectorHelper.ConnectShapes(r1, r2, connector);

            var con_layer = page.Layers["Connector"];

            var sel = page.CreateSelection(
                IVisio.VisSelectionTypes.visSelTypeByLayer,
                IVisio.VisSelectMode.visSelModeSkipSub,
                con_layer);
            sel.SendToBack();
        }
    }
}