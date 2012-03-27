using System;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using TABLEROW = VisioAutomation.ShapeSheet.Data.TableRow<VisioAutomation.ShapeSheet.CellData<double>>;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        // This class is meant for those cell groups that appear as multiple rows in a section
        // for example the character section or the paragraph section

        // descendants must implement this method.
        // the implementation should be this: run the "func" on each formula in the cell
        // group (even in the formula is null) 
        protected abstract void ApplyFormulas(ApplyFormula func, short row);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid, short row)
        {
            this.ApplyFormulas((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f), row);
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            this.ApplyFormulas((src, f) => update.SetFormulaIgnoreNull(src, f),row);
        }

        protected static IList<List<TObj>> CellsFromRowsGrouped<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToCells<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_lists = new List<List<TObj>>(table.Groups.Count);

            for (int group_index = 0; group_index < table.Groups.Count; group_index++)
            {
                var group = table.Groups[group_index];
                var tablerows = group.RowIndices.Select(ri => table[ri]);
                var cells_list = new List<TObj>(group.Count);
                var cells = tablerows.Select(row => row_to_obj_func(query,row));
                cells_list.AddRange(cells);
                list_of_lists.Add(cells_list);
            }

            return list_of_lists;
        }

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToCells<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            var cells = table.Select( row => row_to_obj_func(query, row) );
            var cells_list = new List<TObj>(table.RowCount);
            cells_list.AddRange(cells);
            return cells_list;
        }
    }
}