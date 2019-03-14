using System.Collections.Generic;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderMultiRow<TGroup> where TGroup : CellGroupBase
    {
        protected VASS.Query.SectionsQuery query;

        protected ReaderMultiRow()
        {
            this.query = new VASS.Query.SectionsQuery();
        }
        
        public abstract TGroup ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row);

        public List<List<TGroup>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var data_for_shapes = query.GetCells(page, shapeids, type);

            var list_cellgroups = new List<List<TGroup>>(shapeids.Count);
            foreach (var d in data_for_shapes)
            {
                var first_section = d.Sections[0];
                var cellgroups = this.__ToCellGroups(first_section);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }

        public List<TGroup> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var data_for_shape = query.GetCells(shape,type);
            var first_section = data_for_shape.Sections[0];
            var cellgroups = this.__ToCellGroups(first_section);
            return cellgroups;
        }

        private List<TGroup> __ToCellGroups(VASS.Query.SectionQueryOutput<string> section_data)
        {
            var cellgroups = new List<TGroup>(section_data.Rows.Count);
            foreach (var section_row in section_data.Rows)
            {
                var cellgroup = this.ToCellGroup(section_row.Cells);
                cellgroups.Add(cellgroup);                
            }
            return cellgroups;
        }
    }
}