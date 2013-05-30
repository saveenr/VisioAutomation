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
            //#var doc = docs.Add("");
            short flags = (short)
                (IVisio.VisOpenSaveArgs.visOpenDocked |
                IVisio.VisOpenSaveArgs.visOpenRO);


            var doc = docs.AddEx("orgch_u.vst", IVisio.VisMeasurementSystem.visMSUS, 0,0);
            //var orgchart_stencil = docs.Open("orgch_u.vss");
            var orgchart_masters = doc.Masters;
            var position_master = orgchart_masters["Position Belt"];
            var dyncon = orgchart_masters["Dynamic Connector"];

            var page = app.ActivePage;

            var masters = new object[] {position_master, position_master, dyncon};

            var xy_array = new double[] {1, 2 ,3,5, 0,0};
            System.Array outids_sa;
            page.DropManyU(masters, xy_array, out outids_sa);
            short[] outids = (short[])outids_sa;

            var shape = page.Shapes[outids[0]];

            //return outids;
            //page.DropManyU(masters, 1, 2);
        }
    }
}