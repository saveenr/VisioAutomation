using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetUpdate
    {
        internal readonly ShapeSheet.Update update;
        public Client Client;
        public IVisio.Page TargetPage;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetUpdate(Client client, IVisio.Page page)
        {
            this.Client = client;
            this.TargetPage = page;
            this.update = new ShapeSheet.Update();
        }

        public void SetFormula(short id, ShapeSheet.SRC src, string formula)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update.SetFormula(sidsrc, formula);
        }

        public void SetResult(short id, ShapeSheet.SRC src, double result)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void SetResult(short id, ShapeSheet.SRC src, int result)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void SetResult(short id, ShapeSheet.SRC src, string result)
        {
            var sidsrc = new ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void Update()
        {
            this.Client.WriteVerbose("Staring ShapeSheet Update");
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Update ShapeSheet Formulas"))
            {
                this.update.BlastGuards = this.BlastGuards;
                this.update.TestCircular = this.TestCircular;
                this.update.Execute(this.TargetPage);
            }
            this.Client.WriteVerbose("Ending ShapeSheet Update");
        }
    }
}