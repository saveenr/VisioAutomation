using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T,RT>(IList<VisioAutomation.ShapeSheet.CellData<RT>> data);

        public struct SRCFormulaPair
        {
            public VisioAutomation.ShapeSheet.SRC SRC;
            public VisioAutomation.ShapeSheet.FormulaLiteral Formula;

            public SRCFormulaPair(VisioAutomation.ShapeSheet.SRC src, VisioAutomation.ShapeSheet.FormulaLiteral formula)
            {
                this.SRC = src;
                this.Formula = formula;
            }
        }

        protected SRCFormulaPair newpair(VisioAutomation.ShapeSheet.SRC src, VisioAutomation.ShapeSheet.FormulaLiteral formula)
        {
            return new SRCFormulaPair(src, formula);
        }

        public abstract IEnumerable<SRCFormulaPair> Pairs { get; }
    }
}