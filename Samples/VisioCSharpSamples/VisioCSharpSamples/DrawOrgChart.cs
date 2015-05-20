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

            // Create a new OrgChart using the org chart template
            var doc = docs.AddEx("orgch_u.vst", IVisio.VisMeasurementSystem.visMSUS, 0,0);
            var page = app.ActivePage;

            // Get the masters
            bool is_visio_2013 = app.Version.StartsWith("15.");
            var orgchart_masters = doc.Masters;
            var position_master = orgchart_masters[is_visio_2013 ? "Position Belt" : "Position"];
            var dyncon = orgchart_masters["Dynamic Connector"];

            // three masters to drop
            var masters = new object[] { position_master, position_master, position_master };

            // three coordinates to make the drop
            var xy_array = new double[] {1,2 ,3,5, 6,2};

            // perform the drop of all shapes at once
            // and get the shape ids back out
            System.Array shape_ids_sa;
            page.DropManyU(masters, xy_array, out shape_ids_sa);
            var shape_ids = (short[])shape_ids_sa;

            // using the shape ids, get the shape objects
            var shape0 = page.Shapes.ItemFromID16[shape_ids[0]];
            var shape1 = page.Shapes.ItemFromID16[shape_ids[1]];
            var shape2 = page.Shapes.ItemFromID16[shape_ids[2]];

            // perform connections using the Dynamic Connection master
            shape0.AutoConnect(shape1, IVisio.VisAutoConnectDir.visAutoConnectDirNone, dyncon);
            shape1.AutoConnect(shape2, IVisio.VisAutoConnectDir.visAutoConnectDirNone, dyncon);

        }
    }
}