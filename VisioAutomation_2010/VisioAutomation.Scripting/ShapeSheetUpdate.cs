using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetUpdate
    {
        internal readonly VA.ShapeSheet.Update update;
        public VA.Scripting.Client Client;
        public IVisio.Page TargetPage;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetUpdate(VA.Scripting.Client ss,IVisio.Page page)
        {
            this.Client = ss;
            this.TargetPage = page;
            this.update = new VA.ShapeSheet.Update();
        }

        public void SetFormula(short id, VA.ShapeSheet.SRC src, string formula)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.update.SetFormula(sidsrc, formula);
        }

        public void SetResult(short id, VA.ShapeSheet.SRC src, double result)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void SetResult(short id, VA.ShapeSheet.SRC src, int result)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void SetResult(short id, VA.ShapeSheet.SRC src, string result)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNumber);
        }

        public void Update()
        {
            this.Client.WriteVerbose("Staring ShapeSheet Update");
            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(application, "Update ShapeSheet Formulas"))
            {
                this.update.BlastGuards = this.BlastGuards;
                this.update.TestCircular = this.TestCircular;
                this.update.Execute(this.TargetPage);
            }
            this.Client.WriteVerbose("Ending ShapeSheet Update");
        }
    }
}