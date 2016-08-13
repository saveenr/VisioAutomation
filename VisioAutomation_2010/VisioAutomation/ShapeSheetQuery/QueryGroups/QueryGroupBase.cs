using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.QueryGroups
{
    public abstract class QueryGroupBase
    {
        public delegate T CellsToObject<T,RT>(ShapeSheet.CellData<RT>[] data);

        protected SRCFormulaPair newpair(ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            return new SRCFormulaPair(src, formula);
        }

        public abstract IEnumerable<SRCFormulaPair> Pairs { get; }
    }
}