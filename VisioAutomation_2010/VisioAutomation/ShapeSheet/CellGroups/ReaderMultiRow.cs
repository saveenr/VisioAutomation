using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderMultiRow<TGroup> where TGroup : CellGroupMultiRow
    {
        protected SectionsQuery query;

        protected ReaderMultiRow()
        {
            this.query = new SectionsQuery();
        }
        
        public abstract TGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

        protected void validate_query()
        {
            if (this.query.Queries.Count != 1)
            {
                throw new InternalAssertionException("Query should contain contain exactly 1 subquery");
            }
        }

        public List<List<TGroup>> GetCellGroups(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            this.validate_query();
            var data_for_shapes = query.GetFormulasAndResults(page, shapeids);
            var list = new List<List<TGroup>>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.SubQueryRowsToCellGroups(d.Sections[0]));
            list.AddRange(objects);
            return list;
        }

        public List<TGroup> GetCellGroups(Microsoft.Office.Interop.Visio.Shape shape)
        {
            this.validate_query();
            var data_for_shape = query.GetFormulasAndResults(shape);
            var sec = data_for_shape.Sections[0];
            var cellgroups = this.SubQueryRowsToCellGroups(sec);
            return cellgroups;
        }

        private List<TGroup> SubQueryRowsToCellGroups(SectionQueryOutput<ShapeSheet.CellData> subquery_output)
        {
            var list_celldata = subquery_output.Rows.Select(row => this.CellDataToCellGroup(row.Cells));
            var cellgroups = new List<TGroup>(subquery_output.Rows.Count);
            cellgroups.AddRange(list_celldata);
            return cellgroups;
        }
    }
}