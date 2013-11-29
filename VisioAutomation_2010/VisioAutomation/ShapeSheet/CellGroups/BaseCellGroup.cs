using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate T RowToObject<T>(CellData<double>[] data);
        public delegate T _RowToObject<T>(CellData<string>[] data);
        public delegate T ____RowToObject<T,X>(CellData<X>[] data);

        public delegate T QueryResultToObject<T>(Query.CellQuery.QueryResult<CellData<double>> data);
        public delegate T _QueryResultToObject<T>(Query.CellQuery.QueryResult<CellData<string>> data);
        public delegate T ____QueryResultToObject<T,X>(Query.CellQuery.QueryResult<CellData<X>> data);

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