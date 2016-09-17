using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries.CommonQueries
{
    public abstract class CellGroupQuery<TCellGroup, TResult>
    {
        protected VisioAutomation.ShapeSheet.Queries.Query query;

        protected CellGroupQuery()
        {
            this.query = new Query();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData<TResult>[] row);

    }

    public abstract class CellGroupSingleRowQuery<TCellGroup,TResult>: CellGroupQuery<TCellGroup, TResult>
    {

        protected override void validate_query()
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
            validate_query();

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
            validate_query();
            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = this.query.GetFormulasAndResults<TResult>(surface);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }


    public abstract class CellGroupMultiRowQuery<TCellGroup, TResult> : CellGroupQuery<TCellGroup, TResult>
    {
        protected override void validate_query()
        {
            if (this.query.Cells.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain any cells");
            }

            if (this.query.SubQueries.Count != 1)
            {
                throw new InternalAssertionException("Query should contain contain exactly 1 subquery");
            }
        }

        public List<List<TCellGroup>> GetCells(
            IVisio.Page page,
            IList<int> shapeids)
        {
            this.validate_query();

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetFormulasAndResults<TResult>(surface, shapeids);
            var list = new List<List<TCellGroup>>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.SectionRowsToObjects(d.Sections[0]));
            list.AddRange(objects);
            return list;
        }

        public List<TCellGroup> GetCells(IVisio.Shape shape)
        {
            this.validate_query();
            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = query.GetFormulasAndResults<TResult>(surface);
            var sec = data_for_shape.Sections[0];
            var cellgroups = this.SectionRowsToObjects(sec);
            return cellgroups;
        }

        private List<TCellGroup> SectionRowsToObjects(VisioAutomation.ShapeSheet.Queries.Outputs.SubQueryOutput<ShapeSheet.CellData<TResult>> subquery_output)
        {
            var list_celldata = subquery_output.Rows.Select(row => this.CellDataToCellGroup(row.Cells));
            var cellgroups = new List<TCellGroup>(subquery_output.Rows.Count);
            cellgroups.AddRange(list_celldata);
            return cellgroups;
        }
    }
}