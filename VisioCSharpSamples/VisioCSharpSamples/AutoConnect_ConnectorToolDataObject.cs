using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void AutoConnect_ConnectorToolDataObject(IVisio.Document doc)
        {
            var page = Util.CreateStandardPage(doc, "Using ConnectorToolDataObject");

            var s1 = page.DrawRectangle(0, 0, 1, 1);
            var s2 = page.DrawRectangle(3, 3, 4, 4);

            var app = page.Application;
            var conobj = app.ConnectorToolDataObject;

            s1.AutoConnect(s2,IVisio.VisAutoConnectDir.visAutoConnectDirNone, conobj);
        }
    }
}