using VisioAutomation.ShapeSheet.Writers;

namespace VisioScripting.Models
{
    public class ShapeSheetWriter
    {
        internal readonly SidSrcWriter writer;
        public Client Client;
        public VisioAutomation.SurfaceTarget Surface;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, Microsoft.Office.Interop.Visio.Page page)
        {
            this.Client = client;
            this.Surface = new VisioAutomation.SurfaceTarget(page);
            this.writer = new SidSrcWriter();
        }

        public void SetFormula(short id, VisioAutomation.ShapeSheet.Src src, string formula)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SidSrc(id, src);
            this.writer.SetFormula(sidsrc, formula);
        }

        public void Commit()
        {
            using (var undoscope = this.Client.Application.NewUndoScope(nameof(ShapeSheetWriter)+"."+nameof(Commit)))
            {
                this.writer.BlastGuards = this.BlastGuards;
                this.writer.TestCircular = this.TestCircular;
                this.writer.Commit(this.Surface);
            }
        }
    }
}