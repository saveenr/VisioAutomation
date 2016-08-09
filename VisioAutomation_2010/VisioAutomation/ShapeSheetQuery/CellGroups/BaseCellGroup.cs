using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T,RT>(IList<ShapeSheet.CellData<RT>> data);

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