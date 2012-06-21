using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void Shape_SetFormulas(IVisio.Document doc)
        {
            var page = Util.CreateStandardPage(doc, "SSF");
            var shape = Util.CreateStandardShape(page);

            // CREATE REQUEST
            var request = new[]
                              {
                                  new
                                      {
                                          Section = (short) IVisio.VisSectionIndices.visSectionObject,
                                          Row = (short) IVisio.VisRowIndices.visRowXFormOut,
                                          Cell = (short) IVisio.VisCellIndices.visXFormWidth,
                                          Formula = "2.0"
                                      },
                                  new
                                      {
                                          Section = (short) IVisio.VisSectionIndices.visSectionObject,
                                          Row = (short) IVisio.VisRowIndices.visRowXFormOut,
                                          Cell = (short) IVisio.VisCellIndices.visXFormHeight,
                                          Formula = "3.0"
                                      }
                              };

            // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS

            var SRCStream = new short[request.Length*3];
            var formulas_objects = new object[request.Length];
            for (int i = 0; i < request.Length; i++)
            {
                SRCStream[(i*3) + 0] = request[i].Section;
                SRCStream[(i*3) + 1] = request[i].Row;
                SRCStream[(i*3) + 2] = request[i].Cell;
                formulas_objects[i] = request[i].Formula;
            }

            // EXECUTE THE REQUEST
            short flags = (short) (IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
            int count = shape.SetFormulas(SRCStream, formulas_objects, flags);

            // DISPLAY THE INFORMATION
            shape.Text = "SetFormulas";
        }
    }
}