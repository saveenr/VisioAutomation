namespace VisioAutomation.Scripting.ShapeSheet
{
    public class ShapeSheetWriter
    {
        internal readonly VisioAutomation.ShapeSheet.Writers.FormulaWriterSIDSRC formula_writer;
        public Client Client;
        public VisioAutomation.ShapeSheet.ShapeSheetSurface Surface;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetWriter(Client client, Microsoft.Office.Interop.Visio.Page page)
        {
            this.Client = client;
            this.Surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);
            this.formula_writer = new VisioAutomation.ShapeSheet.Writers.FormulaWriterSIDSRC();
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
                this.formula_writer.Commit(this.Surface);
            }
        }
    }
}