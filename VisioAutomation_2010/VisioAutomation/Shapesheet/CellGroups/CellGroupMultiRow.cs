using System;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        // This class is meant for those cell groups that appear as multiple rows in a section
        // for example the character section or the paragraph section

        // Delegates
        protected delegate TObj RowToObject<TQuery, TObj>(TQuery query, VA.ShapeSheet.Data.TableRow<VA.ShapeSheet.CellData<double>> qds) where TQuery : VA.ShapeSheet.Query.SectionQuery;

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

        protected static IList<List<TObj>> CellsFromRowsGrouped<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToObject<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_lists = new List<List<TObj>>(qds.Groups.Count);

            for (int group_index = 0; group_index < qds.Groups.Count; group_index++)
            {
                var group = qds.Groups[group_index];
                var rows = group.RowIndices.Select(ri => qds[ri]);
                var obj_list = new List<TObj>(group.Count);
                var objs = rows.Select(row => row_to_obj_func(query,row));
                obj_list.AddRange(objs);
                list_of_lists.Add(obj_list);
            }

            return list_of_lists;
        }

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToObject<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            var objs = table.Select( row => row_to_obj_func(query, row) );
            var obj_list = new List<TObj>(table.Count);
            obj_list.AddRange(objs);
            return obj_list;
        }
    }
}