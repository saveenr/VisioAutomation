using System;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupForSection : BaseCellGroup
    {
        // Delegates
        protected delegate TObj RowToObject<TQuery, TObj>(TQuery query, VA.ShapeSheet.Data.QueryDataRow<double> qds) where TQuery : VA.ShapeSheet.Query.SectionQuery;

        protected abstract void ApplyFormulas(ApplyFormula func, short row);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid, short row)
        {
            this.ApplyFormulas((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f), row);
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            this.ApplyFormulas((src, f) => update.SetFormulaIgnoreNull(src, f),row);
        }

        protected static IList<List<TObj>> _GetObjectsFromRowsGrouped<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToObject<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_lists = new List<List<TObj>>(qds.Groups.Count);

            for (int group_index = 0; group_index < qds.Groups.Count; group_index++)
            {
                var group = qds.Groups[group_index];
                var rows = group.RowIndices.Select(ri => qds[ri]);
                var objs = rows.Select(r => row_to_obj_func(query, r));
                var obj_list = new List<TObj>(group.Count);
                obj_list.AddRange(objs);
                list_of_lists.Add(obj_list);
            }

            return list_of_lists;
        }

        protected static IList<TObj> _GetObjectsFromRows<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToObject<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var objs = qds.Select(r => row_to_obj_func(query, r));
            var obj_list = new List<TObj>(qds.Count);
            obj_list.AddRange(objs);
            return obj_list;
        }

    }
}