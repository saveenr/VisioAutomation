using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.ShapeSheet
{
    public abstract class CellDataGroup
    {
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected abstract void _Apply(ApplyFormula func);
        protected delegate TCells row_to_cells<TCells, TQuery>(TQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row) where TQuery : VA.ShapeSheet.Query.CellQuery;

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        protected static IList<TCells> _GetCells<TCells, TQuery>(IVisio.Page page, IList<int> shapeids, TQuery q, row_to_cells<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var cells_list = new List<TCells>(shapeids.Count);
            var qds = q.GetFormulasAndResults<double>(page, shapeids);
            for (int i = 0; i < qds.RowCount; i++)
            {
                var cells = row_to_cells_func(q, qds, i);
                cells_list.Add(cells);
            }

            return cells_list;
        }

        protected static TCells _GetCells<TCells, TQuery>(IVisio.Shape shape, TQuery query, row_to_cells<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            return row_to_cells_func(query, qds, 0);
        }
    }

    public abstract class CellSectionDataGroup
    {
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected abstract void _Apply(ApplyFormula func);
        protected delegate TCells row_to_cells<TCells, TQuery>(TQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row) where TQuery : VA.ShapeSheet.Query.CellQuery;

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        protected static IList<List<TCells>> _GetCells<TCells, TQuery>(IVisio.Page page, IList<int> shapeids, TQuery query, row_to_cells<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var list = new List<List<TCells>>(shapeids.Count);
            foreach (var group in qds.Groups)
            {
                var cells_list = new List<TCells>(group.Count);
                if (group.Count > 0)
                {
                    for (int i = 0; i < qds.RowCount; i++)
                    {
                        var cells = row_to_cells_func(query, qds, i);
                        cells_list.Add(cells);
                    }
                }
                list.Add(cells_list);
            }
            return list;
        }

        protected static IList<TCells> _GetCells<TCells, TQuery>(IVisio.Shape shape, TQuery query, row_to_cells<TCells, TQuery> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var cells_list = new List<TCells>(qds.RowCount);
            for (int row = 0; row < qds.RowCount; row++)
            {
                var cells = row_to_cells_func(query, qds, row);
                cells_list.Add(cells);
            }
            return cells_list;
        }
    }
}