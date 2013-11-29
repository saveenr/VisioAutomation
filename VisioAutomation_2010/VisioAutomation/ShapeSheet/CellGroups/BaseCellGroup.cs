using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T,RT>(CellData<RT>[] data);
        public delegate T QueryResultToObject<T,RT>(Query.CellQuery.QueryResult<CellData<RT>> data);

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