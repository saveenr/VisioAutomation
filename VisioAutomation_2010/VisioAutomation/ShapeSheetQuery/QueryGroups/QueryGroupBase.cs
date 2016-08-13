using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.QueryGroups
{
    public abstract class QueryGroupBase
    {
        public delegate T CellsToObject<T,RT>(ShapeSheet.CellData<RT>[] data);

        public struct SRCFormulaPair
        {
            public ShapeSheet.SRC SRC;
            public ShapeSheet.FormulaLiteral Formula;

            public SRCFormulaPair(ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
            {
                this.SRC = src;
                this.Formula = formula;
            }
        }

        protected SRCFormulaPair newpair(ShapeSheet.SRC src, ShapeSheet.FormulaLiteral formula)
        {
            return new SRCFormulaPair(src, formula);
        }

        public abstract IEnumerable<SRCFormulaPair> Pairs { get; }
    }
}