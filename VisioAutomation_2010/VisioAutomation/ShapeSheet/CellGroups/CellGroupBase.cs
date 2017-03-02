using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupBase
    {
        protected SRCFormulaPair newpair(ShapeSheet.Src src, ShapeSheet.CellValueLiteral formula)
        {
            return new SRCFormulaPair(src, formula);
        }

        public abstract IEnumerable<SRCFormulaPair> SRCFormulaPairs { get; }
    }
}