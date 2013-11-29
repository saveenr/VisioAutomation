using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T,X>(CellData<X>[] data);
        public delegate T QueryResultToObject<T,X>(Query.CellQuery.QueryResult<CellData<X>> data);

        public struct SRCValuePair
        {
            public SRC SRC;
            public FormulaLiteral Formula;

            public SRCValuePair(SRC src, FormulaLiteral f)
            {
                this.SRC = src;
                this.Formula = f;
            }
        }

        protected SRCValuePair createpair(SRC src, FormulaLiteral f)
        {
            return new SRCValuePair(src, f);
        }

        public abstract IEnumerable<SRCValuePair> EnumPairs();
    }
}