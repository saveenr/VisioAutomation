using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries.CommonQueries
{
    public abstract class CellGroupSingleRowQuery<TCellGroup,TResult>
    {
        protected VisioAutomation.ShapeSheet.Queries.Query query;

        protected CellGroupSingleRowQuery()
        {
            this.query = new Query();
        }

        void verify_singlerow_query()
        {
            if (this.query.Cells.Count < 1)
            {
                throw new InternalAssertionException("Query must contain at least one cell");
            }

            if (this.query.SubQueries.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain contain any subqueries");
            }
        }

        public List<TCellGroup> GetCells(
            Microsoft.Office.Interop.Visio.Page page,
            IList<int> shapeids)
        {
            verify_singlerow_query();

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = this.query.GetFormulasAndResults<TResult>(surface, shapeids);
            var list = new List<TCellGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TCellGroup GetCells(
            Microsoft.Office.Interop.Visio.Shape shape)
        {
            verify_singlerow_query();

            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = this.query.GetFormulasAndResults<TResult>(surface);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData<TResult>[] row);
    }


    public abstract class CellGroupMultiRowQuery<TCellGroup, TResult>
    {
        protected VisioAutomation.ShapeSheet.Queries.Query query;

        protected CellGroupMultiRowQuery()
        {
            this.query = new Query();
        }



        private void verify_multirow_query(Query query)
        {
            if (query.Cells.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain any cells");
            }

            if (query.SubQueries.Count != 1)
            {
                throw new InternalAssertionException("Query should contain contain exactly 1 subquery");
            }
        }


        public List<List<TCellGroup>> GetCells(
            IVisio.Page page,
            IList<int> shapeids)
        {
            this.verify_multirow_query(query);

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetFormulasAndResults<TResult>(surface, shapeids);
            var list = new List<List<TCellGroup>>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.SectionRowsToObjects(d.Sections[0]));
            list.AddRange(objects);
            return list;
        }

        public List<TCellGroup> GetCells(
            IVisio.Shape shape)
        {
            this.verify_multirow_query(query);

            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = query.GetFormulasAndResults<TResult>(surface);
            var sec = data_for_shape.Sections[0];
            var sec_objects = this.SectionRowsToObjects(sec);

            return sec_objects;
        }

        private List<TCellGroup> SectionRowsToObjects(
            VisioAutomation.ShapeSheet.Queries.Outputs.SubQueryOutput<ShapeSheet.CellData<TResult>> sec)
        {
            var sec_objects = new List<TCellGroup>(sec.Rows.Count);
            var objects = sec.Rows.Select(row => this.CellDataToCellGroup(row.Cells));
            sec_objects.AddRange(objects);
            return sec_objects;
        }

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData<TResult>[] row);

    }

}