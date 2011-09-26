using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public abstract class CellSectionDataGroup
    {
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected abstract void _Apply(ApplyFormula func, short row);
        protected delegate TCells row_to_cells<TCells, TQuery>(TQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row) where TQuery : VA.ShapeSheet.Query.SectionQuery;
        protected delegate TCells row_to_cells2<TCells, TQuery>(TQuery query, VA.ShapeSheet.Query.QueryDataRow<double> qds) where TQuery : VA.ShapeSheet.Query.SectionQuery;

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f), row);
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f),row);
        }

        protected static IList<List<TCells>> _GetCells<TCells, TQuery>(IVisio.Page page, IList<int> shapeids, TQuery query, row_to_cells<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_lists = new List<List<TCells>>(shapeids.Count);
            foreach (var group in qds.Groups)
            {
                var objs = new List<TCells>(group.Count);
                if (group.Count > 0)
                {
                    for (int i = group.StartRow; i <= group.EndRow; i++)
                    {
                        var obj = row_to_cells_func(query, qds, i);
                        objs.Add(obj);
                    }
                }
                list_of_lists.Add(objs);
            }
            return list_of_lists;
        }

        protected static IList<TCells> _GetCells<TCells, TQuery>(IVisio.Shape shape, TQuery query, row_to_cells<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var objs = new List<TCells>(qds.RowCount);

            for (int row = 0; row < qds.RowCount; row++)
            {
                var obj = row_to_cells_func(query, qds, row);
                objs.Add(obj);
            }
            return objs;
        }

        protected static IList<TCells> _GetCells<TCells, TQuery>(IVisio.Shape shape, TQuery query, row_to_cells2<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var objs = new List<TCells>(qds.RowCount);
            foreach (var row in qds.EnumRows())
            {
                var obj = row_to_cells_func(query, row);
                objs.Add(obj);
                
            }
            return objs;
        }

    }
}