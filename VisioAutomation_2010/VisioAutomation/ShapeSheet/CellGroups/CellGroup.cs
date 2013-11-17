using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        protected static IList<T> _GetCells<T>(IVisio.Page page, IList<int> shapeids, VA.ShapeSheet.Query.CellQuery cellQuery, QueryResultToObject<T> f)
        {
            var data_for_shapes = cellQuery.GetFormulasAndResults<double>(page, shapeids);
            var list = new List<T>(shapeids.Count);
            foreach (var data_for_shape in data_for_shapes)
            {
                var cells = f(data_for_shape);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T>(IVisio.Shape shape, VA.ShapeSheet.Query.CellQuery cellQuery, QueryResultToObject<T> f)
        {
            var data_for_shape = cellQuery.GetFormulasAndResults<double>(shape);
            var cells = f(data_for_shape);
            return cells;
        }

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

        protected SRCValuePair foo(SRC src, FormulaLiteral f)
        {
            return new SRCValuePair(src, f);
        }

        public void ApplyFormulas(ApplyFormula func)
        {
            foreach (var pair in this.EnumPairs())
            {
                func(pair.SRC, pair.Formula);
            }
        }

        public abstract IEnumerable<VA.ShapeSheet.CellGroups.CellGroup.SRCValuePair> EnumPairs();
    }
}