using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Utilities;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public List<VisioAutomation.ShapeSheet.SIDSRC> SIDSRCs;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new ShapeSheetSurface(page);
            this.SIDSRCs = new List<VisioAutomation.ShapeSheet.SIDSRC>();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.SRC src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.SIDSRCs.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var stream = StreamBuilderSIDSRC.CreateStream(this.SIDSRCs);
            var formulas = this.Surface.GetFormulasU(stream);
            return formulas;
        }

        public string[] GetResults()
        {
            var stream = StreamBuilderSIDSRC.CreateStream(this.SIDSRCs);
            var unitcodes = new List<IVisio.VisUnitCodes> { IVisio.VisUnitCodes.visNoCast };
            var formulas = this.Surface.GetResults<string>(stream, unitcodes);
            return formulas;
        }
    }
}