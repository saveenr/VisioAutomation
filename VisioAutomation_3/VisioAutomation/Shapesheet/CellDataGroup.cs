using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.ShapeSheet
{
    public abstract class CellDataGroup
    {
        public delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        protected abstract void _Apply(ApplyFormula func);
    }
}