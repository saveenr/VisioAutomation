using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        // This class is meant for those cell groups that appear as multiple rows in a section
        // for example the character section or the paragraph section

        public abstract void ApplyFormulasForRow(ApplyFormula func, short row);

        protected static IList<List<TObj>> CellsFromRowsGrouped<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToCells<TQuery, TObj> row_to_obj) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(page, shapeids);
            var list_of_groups = new List<List<TObj>>(table.Groups.Count);

            for (int group_index = 0; group_index < table.Groups.Count; group_index++)
            {
                var group = table.Groups[group_index];
                var objects = new List<TObj>(group.Count);
                foreach (int i in group.RowIndices)
                {
                    var new_object = row_to_obj(query, table, i);
                    objects.Add(new_object);
                }
                list_of_groups.Add(objects);
            }

            return list_of_groups;
        }

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToCells<TQuery, TObj> row_to_obj) where TQuery : VA.ShapeSheet.Query.SectionQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            var objects = new List<TObj>(table.RowCount);
            for (int i = 0; i < table.RowCount; i++)
            {
                var new_object = row_to_obj(query, table, i);
                objects.Add(new_object);
            }
            return objects;
        }
    }

    public abstract class CellGroupMultiRowEx : CellGroupMultiRow
    {

        public static IList<List<T>> _GetCells<T>(IVisio.Page page, IList<int> shapeids, QueryEx query, System.Func<CellData<double>[],T> f)
        {

            var outer_list = new List<List<T>>();

 
            var data_for_shapes = query.GetFormulasAndResults<double>(page, shapeids);

            foreach (var data_for_shape in data_for_shapes)
            {
                var inner_list = new List<T>();
                outer_list.Add(inner_list);

                var sec = data_for_shape.SectionCells[0];
                foreach (var row in sec.Rows)
                {
                    var cells = f(row);
                    inner_list.Add(cells);
                }

            }

            return outer_list;
        }


        public static IList<T> _GetCells<T>(IVisio.Shape shape, QueryEx query, System.Func<CellData<double>[], T> f)
        {


            var data_for_shape = query.GetFormulasAndResults<double>(shape);

            var inner_list = new List<T>();

            var sec = data_for_shape.SectionCells[0];
            foreach (var row in sec.Rows)
            {
                var cells = f(row);
                inner_list.Add(cells);
            }

            return inner_list;
        }
    }
}