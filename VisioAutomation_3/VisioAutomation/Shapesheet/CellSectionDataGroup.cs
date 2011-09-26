using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public abstract class CellSectionDataGroup
    {
        // Delegates
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected delegate TObj RowToObject<TObj, TQuery>(TQuery query, VA.ShapeSheet.Query.QueryDataRow<double> qds) where TQuery : VA.ShapeSheet.Query.SectionQuery;

        protected abstract void _Apply(ApplyFormula func, short row);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f), row);
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f),row);
        }

        protected static IList<List<TObj>> _GetObjectsFromRowsGrouped<TObj, TQuery>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToObject<TObj, TQuery> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_lists = new List<List<TObj>>(shapeids.Count);

            for (int group_index = 0; group_index < qds.Groups.Count; group_index++)
            {
                var group = qds.Groups[group_index];
                var rows_in_group = qds.EnumRowsInGroup(group_index);

                var obj_list = new List<TObj>(group.Count);
                foreach (var row in rows_in_group)
                {
                    var obj = row_to_obj_func(query, row);
                    obj_list.Add(obj);
                }

                list_of_lists.Add(obj_list);
            }

            return list_of_lists;
        }

        protected static IList<TObj> _GetObjectsFromRows<TObj, TQuery>(IVisio.Shape shape, TQuery query, RowToObject<TObj, TQuery> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var rows_in_group = qds.EnumRows();

            var obj_list = new List<TObj>(qds.RowCount);
            foreach (var row in rows_in_group)
            {
                var obj = row_to_obj_func(query, row);
                obj_list.Add(obj);
            }

            return obj_list;
        }

    }
}