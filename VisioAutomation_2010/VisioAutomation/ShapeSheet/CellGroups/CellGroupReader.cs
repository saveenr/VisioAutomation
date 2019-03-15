using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public enum CellGroupReaderType
    {
        SingleRow,
        MultiRow
    }

    public abstract class CellGroupReader<TGroup> where TGroup : CellGroup, new()
    {
        protected Query.CellQuery query_singlerow;
        protected VASS.Query.SectionsQuery query_multirow;

        private CellGroupReader()
        {
            this.query_singlerow = null;
            this.query_multirow = null;
        }

        protected CellGroupReader(CellGroupReaderType type)
        {
            if (type == CellGroupReaderType.SingleRow)
            {
                this.query_singlerow = new Query.CellQuery();
                this.query_multirow = null;
            }
            else
            {
                this.query_singlerow = null;
                this.query_multirow = new Query.SectionsQuery();
            }


            var temp_cells = new TGroup();
            if (this.query_singlerow != null)
            {
                foreach (var pair in temp_cells.CellMetadata)
                {
                    this.query_singlerow.Columns.Add(pair.Src, pair.Name);
                }
            }
            else if (this.query_multirow != null)
            {
                var first_cell_metadata = temp_cells.CellMetadata.First();
                var sec = this.query_multirow.SectionQueries.Add((IVisio.VisSectionIndices)first_cell_metadata.Src.Section);
                foreach (var pair in temp_cells.CellMetadata)
                {
                    sec.Columns.Add(pair.Src, pair.Name);
                }
            }
            else
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }
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