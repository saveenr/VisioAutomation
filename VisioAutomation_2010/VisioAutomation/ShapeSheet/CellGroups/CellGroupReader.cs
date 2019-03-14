using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupReader<TGroup> where TGroup : CellGroup
    {
        protected Query.CellQuery query_singlerow;
        protected VASS.Query.SectionsQuery query_multirow;

        private CellGroupReader()
        {
            this.query_singlerow = null;
            this.query_multirow = null;
        }

        protected CellGroupReader(Query.CellQuery query)
        {
            this.query_singlerow = query;
            this.query_multirow = null;
        }

        protected CellGroupReader(Query.SectionsQuery query)
        {
            this.query_singlerow = null;
            this.query_multirow = query;
        }

        public abstract TGroup ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row);

        public List<TGroup> GetCellsSingleRow(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var data_for_shapes = this.query_singlerow.GetCells(page, shapeids, type);
            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.ToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TGroup GetCellsSingleRow(IVisio.Shape shape, CellValueType type)
        {
            var data_for_shape = this.query_singlerow.GetCells(shape, type);
            var cells = this.ToCellGroup(data_for_shape.Cells);
            return cells;
        }


        public List<List<TGroup>> GetCellsMultiRow(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var data_for_shapes = query_multirow.GetCells(page, shapeids, type);

            var list_cellgroups = new List<List<TGroup>>(shapeids.Count);
            foreach (var d in data_for_shapes)
            {
                var first_section = d.Sections[0];
                var cellgroups = this.__ToCellGroups(first_section);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }

        public List<TGroup> GetCellsMultiRow(IVisio.Shape shape, CellValueType type)
        {
            var data_for_shape = query_multirow.GetCells(shape, type);
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