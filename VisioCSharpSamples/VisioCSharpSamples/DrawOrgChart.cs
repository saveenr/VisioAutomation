using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void DrawOrgChart()
        {
            var app = new IVisio.ApplicationClass();
            var docs = app.Documents;

            var doc = docs.AddEx("orgch_u.vst", IVisio.VisMeasurementSystem.visMSUS, 0,0);
            var orgchart_masters = doc.Masters;

            var position_master_name = app.Version.StartsWith( "15.") ? "Position Belt" : "Position";

            var position_master = orgchart_masters[position_master_name];
            var dyncon = orgchart_masters["Dynamic Connector"];

            var page = app.ActivePage;

            var masters = new object[] {position_master, position_master, position_master };

            var xy_array = new double[] {1,2 ,3,5, 6,7};

            System.Array shape_ids_sa;
            page.DropManyU(masters, xy_array, out shape_ids_sa);
            short[] shape_ids = (short[])shape_ids_sa;

            var shape0 = page.Shapes.ItemFromID16[shape_ids[0]];
            var shape1 = page.Shapes.ItemFromID16[shape_ids[1]];
            var shape2 = page.Shapes.ItemFromID16[shape_ids[2]];
            shape0.AutoConnect(shape1, IVisio.VisAutoConnectDir.visAutoConnectDirNone, dyncon);
            shape1.AutoConnect(shape2, IVisio.VisAutoConnectDir.visAutoConnectDirNone, dyncon);

        }
    }
}