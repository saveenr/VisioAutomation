using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T,RT>(IList<CellData<RT>> data);

        public struct SRCFormulaPair
        {
            public SRC SRC;
            public FormulaLiteral Formula;

            public SRCFormulaPair(SRC src, FormulaLiteral formula)
            {
                this.SRC = src;
                this.Formula = formula;
            }
        }

        protected SRCFormulaPair newpair(SRC src, FormulaLiteral formula)
        {
            return new SRCFormulaPair(src, formula);
        }

        public abstract IEnumerable<SRCFormulaPair> Pairs { get; }
    }
}