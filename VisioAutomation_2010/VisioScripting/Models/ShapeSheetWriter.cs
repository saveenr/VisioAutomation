using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class ShapeSheetWriter
    {
        internal readonly SidSrcWriter Writer;
        public Client Client;
        public VisioAutomation.SurfaceTarget Surface;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.Surface = new VisioAutomation.SurfaceTarget(page);
            this.Writer = new SidSrcWriter();
        }

        public void SetFormula(short id, VisioAutomation.ShapeSheet.Src src, string formula)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SidSrc(id, src);
            this.Writer.SetValue(sidsrc, formula);
        }

        public void Commit()
        {
            using (var undoscope = this.Client.Undo.NewUndoScope(nameof(ShapeSheetWriter)+"."+nameof(Commit)))
            {
                this.Writer.BlastGuards = this.BlastGuards;
                this.Writer.TestCircular = this.TestCircular;
                this.Writer.CommitFormulas(this.Surface);
            }
        }
    }
}