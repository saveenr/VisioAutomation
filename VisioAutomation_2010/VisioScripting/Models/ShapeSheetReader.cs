

namespace VisioScripting.Models
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.SurfaceTarget Surface;
        public List<VisioAutomation.ShapeSheet.SidSrc> SidSrcs;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new VisioAutomation.SurfaceTarget(page);
            this.SidSrcs = new List<VisioAutomation.ShapeSheet.SidSrc>();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.Src src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SidSrc(id, src);
            this.SidSrcs.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var stream = VisioAutomation.ShapeSheet.Streams.StreamArray.FromSidSrc(this.SidSrcs);
            var formulas = this.Surface.GetFormulasU(stream);
            return formulas;
        }

        public string[] GetResults()
        {
            const object [] unitcodes = null;
            var stream = VisioAutomation.ShapeSheet.Streams.StreamArray.FromSidSrc(this.SidSrcs);
            var formulas = this.Surface.GetResults<string>( stream, unitcodes);
            return formulas;
        }
    }
}