using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public VisioAutomation.ShapeSheet.Query.SIDSRCStream SIDSRCStream;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new ShapeSheetSurface(page);
            this.SIDSRCStream = new SIDSRCStream();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.SRC src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.SIDSRCStream.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var formulas = this.Surface.GetFormulasU(this.SIDSRCStream);
            return formulas;
        }

        public string[] GetResults()
        {
            var unitcodes = new List<IVisio.VisUnitCodes> { IVisio.VisUnitCodes.visNoCast };
            var formulas = this.Surface.GetResults<string>( this.SIDSRCStream, unitcodes);
            return formulas;
        }
    }
}