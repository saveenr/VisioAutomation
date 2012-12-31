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

        public abstract void ApplyFormulasForRow(ApplyFormula func, short row);

        protected static IList<List<TObj>> CellsFromRowsGrouped<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToCells<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_lists = new List<List<TObj>>(table.Groups.Count);

            for (int group_index = 0; group_index < table.Groups.Count; group_index++)
            {
                var group = table.Groups[group_index];
                var cells_list = new List<TObj>(group.Count);
                foreach (int i in group.RowIndices)
                {
                    var new_object = row_to_obj_func(query, table, i);
                    cells_list.Add(new_object);
                }
                list_of_lists.Add(cells_list);
            }

            return list_of_lists;
        }

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToCells<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            var cells_list = new List<TObj>(table.RowCount);
            for (int i = 0; i < table.RowCount; i++)
            {
                var new_object = row_to_obj_func(query, table, i);
                cells_list.Add(new_object);
            }
            return cells_list;
        }
    }
}