using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void Shape_SetResults(IVisio.Document doc)
        {
            var page = Util.CreateStandardPage(doc, "SSR");
            var shape = Util.CreateStandardShape(page);

            // CREATE REQUEST
            var request = new[]
                              {
                                  new
                                      {
                                          Section = (short) IVisio.VisSectionIndices.visSectionObject,
                                          Row = (short) IVisio.VisRowIndices.visRowXFormOut,
                                          Cell = (short) IVisio.VisCellIndices.visXFormWidth,
                                          UnitCode = (short) IVisio.VisUnitCodes.visNoCast,
                                          Result = (double) 8.2
                                      },
                                  new
                                      {
                                          Section = (short) IVisio.VisSectionIndices.visSectionObject,
                                          Row = (short) IVisio.VisRowIndices.visRowXFormOut,
                                          Cell = (short) IVisio.VisCellIndices.visXFormHeight,
                                          UnitCode = (short) IVisio.VisUnitCodes.visNoCast,
                                          Result = (double) 1.3
                                      }
                              };

            // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
            var SRCStream = new short[request.Length*3];
            var results_objects = new object[request.Length];
            var unitcodes = new object[request.Length];
            for (int i = 0; i < request.Length; i++)
            {
                SRCStream[(i*3) + 0] = request[i].Section;
                SRCStream[(i*3) + 1] = request[i].Row;
                SRCStream[(i*3) + 2] = request[i].Cell;
                results_objects[i] = request[i].Result;
                unitcodes[i] = request[i].UnitCode;
            }

            // EXECUTE THE REQUEST
            short flags = 0;
            int count = shape.SetResults(SRCStream, unitcodes, results_objects, flags);

            // DISPLAY THE INFORMATION
            shape.Text = "SetResults";
        }
    }
}