using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetWriter
    {
        internal readonly FormulaWriterSIDSRC formula_writer;
        public Client Client;
        public IVisio.Page TargetPage;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.TargetPage = page;
            this.formula_writer = new FormulaWriterSIDSRC();
        }

        public void SetFormula(short id, VisioAutomation.ShapeSheet.SRC src, string formula)
        {
            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC(id, src);
            this.formula_writer.SetFormula(sidsrc, formula);
        }

        public void Commit()
        {
            using (var undoscope = this.Client.Application.NewUndoScope("Modify ShapeSheet"))
            {
                this.formula_writer.BlastGuards = this.BlastGuards;
                this.formula_writer.TestCircular = this.TestCircular;
                this.formula_writer.Commit(this.TargetPage);
            }
        }
    }
}