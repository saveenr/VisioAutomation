using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class ShapeSheetUpdate
    {
        internal readonly VA.ShapeSheet.Update update;

        public ShapeSheetUpdate()
        {
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

    }
}