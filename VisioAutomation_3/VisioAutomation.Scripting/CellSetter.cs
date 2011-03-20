using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class CellSetter
    {
        internal readonly VA.ShapeSheet.Update.SIDSRCUpdate update;

        public CellSetter()
        {
            this.update = new VA.ShapeSheet.Update.SIDSRCUpdate();
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
    }
}