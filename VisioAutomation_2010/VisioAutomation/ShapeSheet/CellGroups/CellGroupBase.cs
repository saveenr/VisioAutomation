using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupBase
    {
        protected SrcFormulaPair newpair(ShapeSheet.Src src, ShapeSheet.CellValueLiteral formula)
        {
            return new SrcFormulaPair(src, formula);
        }

        public abstract IEnumerable<SrcFormulaPair> SrcFormulaPairs { get; }
    }
}