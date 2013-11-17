using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        public delegate T RowToObject<T>(CellData<double>[] data);
        public delegate T QueryResultToObject<T>(VA.ShapeSheet.Query.CellQuery.QueryResult<CellData<double>>  data);

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

        public abstract IEnumerable<VA.ShapeSheet.CellGroups.BaseCellGroup.SRCValuePair> EnumPairs();

        public void ApplyFormulas(ApplyFormula func)
        {
            foreach (var pair in this.EnumPairs())
            {
                func(pair.SRC, pair.Formula);
            }
        }

        public void ApplyFormulas(ApplyFormula func, short row)
        {
            foreach (var pair in this.EnumPairs())
            {
                func(pair.SRC.ForRow(row), pair.Formula);
            }
        }
    }
}