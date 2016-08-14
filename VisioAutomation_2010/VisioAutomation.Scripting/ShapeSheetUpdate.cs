using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetUpdate
    {
        internal readonly FormulaWriterSIDSRC update;
        internal readonly ResultWriterSIDSRC  update2;
        public Client Client;
        public IVisio.Page TargetPage;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetUpdate(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.TargetPage = page;
            this.update = new FormulaWriterSIDSRC();
            this.update2 = new ResultWriterSIDSRC();
        }

        public void SetFormula(short id, ShapeSheet.SRC src, string formula)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update.SetFormula(sidsrc, formula);
        }

        public void SetResult(short id, ShapeSheet.SRC src, double result)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update2.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void SetResult(short id, ShapeSheet.SRC src, int result)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update2.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void SetResult(short id, ShapeSheet.SRC src, string result)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update2.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void Update()
        {
            this.Client.WriteVerbose("Staring ShapeSheet Update");
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Update ShapeSheet Formulas"))
            {
                this.update.BlastGuards = this.BlastGuards;
                this.update.TestCircular = this.TestCircular;
                this.update.Commit(this.TargetPage);

                this.update2.BlastGuards = this.BlastGuards;
                this.update2.TestCircular = this.TestCircular;
                this.update2.Commit(this.TargetPage);

            }
            this.Client.WriteVerbose("Ending ShapeSheet Update");
        }
    }
}