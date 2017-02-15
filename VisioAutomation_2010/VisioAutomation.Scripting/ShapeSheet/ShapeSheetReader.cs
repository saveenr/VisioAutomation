using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public SIDSRCStreamBuilder SidsrcStreamBuilder;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new ShapeSheetSurface(page);
            this.SidsrcStreamBuilder = new SIDSRCStreamBuilder();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.SRC src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.SidsrcStreamBuilder.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var formulas = this.Surface.GetFormulasU(this.SidsrcStreamBuilder.ToStream());
            return formulas;
        }

        public string[] GetResults()
        {
            var builder = new ShapeSheetObjectArrayBuilder<IVisio.VisUnitCodes>(1);
            builder.Add(IVisio.VisUnitCodes.visNoCast);

            var unitcodes = builder.ToObjectArray();
            var formulas = this.Surface.GetResults<string>( this.SidsrcStreamBuilder.ToStream(), unitcodes);
            return formulas;
        }
    }
}