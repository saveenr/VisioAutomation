using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetReader
    {
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public VisioAutomation.ShapeSheet.Streams.SIDSRCStreamBuilder SidsrcStreamBuilder;
        
        public ShapeSheetReader(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new ShapeSheetSurface(page);
            this.SidsrcStreamBuilder = new VisioAutomation.ShapeSheet.Streams.SIDSRCStreamBuilder();
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
            const object [] unitcodes = null;
            var stream = this.SidsrcStreamBuilder.ToStream();
            var formulas = this.Surface.GetResults<string>( stream, unitcodes);
            return formulas;
        }
    }
}