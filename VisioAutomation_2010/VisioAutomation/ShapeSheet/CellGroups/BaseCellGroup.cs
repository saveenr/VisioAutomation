using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T,RT>(CellData<RT>[] data);

        public struct SRCValuePair
        {
            public SRC SRC;
            public FormulaLiteral Formula;

            public SRCValuePair(SRC src, FormulaLiteral formula)
            {
                this.SRC = src;
                this.Formula = formula;
            }
        }

        protected SRCValuePair srcvaluepair(SRC src, FormulaLiteral f)
        {
            return new SRCValuePair(src, f);
        }

        public abstract IEnumerable<SRCValuePair> EnumPairs();
    }
}