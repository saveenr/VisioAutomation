using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetUpdate
    {
        internal readonly VA.ShapeSheet.Update update;
        public VA.Scripting.Session ScriptinSession;
        public IVisio.Page TargetPage;
        public bool BlastGuards;
        public bool TestCircular;

        public ShapeSheetUpdate(VA.Scripting.Session ss,IVisio.Page page)
        {
            this.ScriptinSession = ss;
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
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNoCast);
        }

        public void SetResult(short id, VA.ShapeSheet.SRC src, int result)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNoCast);
        }

        public void SetResult(short id, VA.ShapeSheet.SRC src, string result)
        {
            var sidsrc = new VA.ShapeSheet.SIDSRC(id, src);
            this.update.SetResult(sidsrc, result, IVisio.VisUnitCodes.visNoCast);
        }

        public void Update()
        {
            this.ScriptinSession.WriteVerbose("Staring ShapeSheet Update");
            var application = this.ScriptinSession.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(application, "Update ShapeSheet Formulas"))
            {
                this.update.BlastGuards = this.BlastGuards;
                this.update.TestCircular = this.TestCircular;
                this.update.Execute(this.TargetPage);
            }
            this.ScriptinSession.WriteVerbose("Ending ShapeSheet Update");
        }
    }
}