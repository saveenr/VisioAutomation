using System.Collections.Generic;
using System.Linq;
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
        
        public abstract TGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row);

        public List<List<TGroup>> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var data_for_shapes = query.GetFormulas(page, shapeids);

            var list_cellgroups = new List<List<TGroup>>(shapeids.Count);
            foreach (var d in data_for_shapes)
            {
                var first_section = d.Sections[0];
                var cellgroups = this.__SectionRowsToCellGroups(first_section);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }

        public List<List<TGroup>> GetResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var data_for_shapes = query.GetResults<string>(page, shapeids);

            var list_cellgroups = new List<List<TGroup>>(shapeids.Count);
            foreach (var d in data_for_shapes)
            {
                var first_section = d.Sections[0];
                var cellgroups = this.__SectionRowsToCellGroups(first_section);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }


        public List<TGroup> GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var data_for_shape = query.GetFormulas(shape);
            var first_section = data_for_shape.Sections[0];
            var cellgroups = this.__SectionRowsToCellGroups(first_section);
            return cellgroups;
        }

        public List<TGroup> GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var data_for_shape = query.GetResults<string>(shape);
            var first_section = data_for_shape.Sections[0];
            var cellgroups = this.__SectionRowsToCellGroups(first_section);
            return cellgroups;
        }


        private List<TGroup> __SectionRowsToCellGroups(SectionQueryOutput<string> section_data)
        {
            var cellgroups = new List<TGroup>(section_data.Rows.Count);
            foreach (var section_row in section_data.Rows)
            {
                var cellgroup = this.CellDataToCellGroup(section_row.Cells);
                cellgroups.Add(cellgroup);                
            }
            return cellgroups;
        }
    }
}