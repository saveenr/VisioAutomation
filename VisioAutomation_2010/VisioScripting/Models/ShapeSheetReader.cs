using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.SurfaceTarget Surface;
        public VisioAutomation.ShapeSheet.Streams.SidSrcStreamBuilder SidSrcStreamBuilder;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new VisioAutomation.SurfaceTarget(page);
            this.SidSrcStreamBuilder = new VisioAutomation.ShapeSheet.Streams.SidSrcStreamBuilder();
        }

        public void AddCell(short id, VisioAutomation.ShapeSheet.Src src)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SidSrc(id, src);
            this.SidSrcStreamBuilder.Add(sidsrc);
        }

        public string[] GetFormulas()
        {
            var formulas = this.Surface.GetFormulasU(this.SidSrcStreamBuilder.ToStream());
            return formulas;
        }

        public string[] GetResults()
        {
            const object [] unitcodes = null;
            var stream = this.SidSrcStreamBuilder.ToStream();
            var formulas = this.Surface.GetResults<string>( stream, unitcodes);
            return formulas;
        }
    }
}