using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupBuilder<TGroup> where TGroup : CellGroup, new()
    {
        public readonly CellGroupBuilderType Type;
        protected Query.CellQuery query_cells_singlerow;
        protected VASS.Query.SingleSectionQuery query_sections_multirow;

        private CellGroupBuilder()
        {
            this.query_cells_singlerow = null;
            this.query_sections_multirow = null;
        }

        protected CellGroupBuilder(CellGroupBuilderType type)
        {
            var temp_cells = new TGroup();
            Query.ColumnList cols;

            this.Type = type;
            if (type == CellGroupBuilderType.SingleRow)
            {
                this.query_cells_singlerow = new Query.CellQuery();
                cols = this.query_cells_singlerow.Columns;
            }
            else if (type == CellGroupBuilderType.MultiRow)
            {
                this.query_sections_multirow = new Query.SingleSectionQuery();
                var query_section = this.query_sections_multirow.SectionQueries.Add(temp_cells.CellMetadata.First().Src);
                cols = query_section.Columns;
            }
            else
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            foreach (var pair in temp_cells.CellMetadata)
            {
                cols.Add(pair.Src, pair.Name);
            }

        }

        public abstract TGroup ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols);

        public List<TGroup> GetCellsSingleRow(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.SingleRow);
            var data_for_shapes = this.GetCells(query_cells_singlerow, page, shapeids, type);
            var list = new List<TGroup>(shapeids.Count);
            var cols = this.query_cells_singlerow.Columns;
            var objects = data_for_shapes.Select(d => this.ToCellGroup(d.Cells,cols));
            list.AddRange(objects);
            return list;
        }

        private void EnforceType(CellGroupBuilderType t)
        {
            if (this.Type != t)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }
        }

        public TGroup GetCellsSingleRow(IVisio.Shape shape, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.SingleRow);
            var data_for_shape = this.GetCells(query_cells_singlerow, shape, type);
            var cols = this.query_cells_singlerow.Columns;
            var cells = this.ToCellGroup(data_for_shape.Cells,cols);
            return cells;
        }
        
        public List<List<TGroup>> GetCellsMultiRow(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.MultiRow);
            var cols = this.query_sections_multirow.SectionQueries[0].Columns;

            var data_for_shapes = GetCells(query_sections_multirow,page, shapeids, type);
            var list_cellgroups = new List<List<TGroup>>(shapeids.Count);
            foreach (var data_for_shape in data_for_shapes)
            {
                var first_section = data_for_shape[0];
                var cellgroups = this.__ToCellGroups(first_section,cols);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }

        public List<TGroup> GetCellsMultiRow(IVisio.Shape shape, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.MultiRow);
            var cols = this.query_sections_multirow.SectionQueries[0].Columns;
            var data_for_shape = GetCells(query_sections_multirow, shape, type);
            var first_section = data_for_shape[0];
            var cellgroups = this.__ToCellGroups(first_section,cols);
            return cellgroups;
        }

        private List<TGroup> __ToCellGroups(VASS.Query.ShapeSectionOutput<string> section_data, VisioAutomation.ShapeSheet.Query.ColumnList cols)
        {
            var cellgroups = new List<TGroup>(section_data.Rows.Count);
            foreach (var section_row in section_data.Rows)
            {
                var cellgroup = this.ToCellGroup(section_row.Cells,cols);
                cellgroups.Add(cellgroup);
            }
            return cellgroups;
        }

        private VASS.Query.ShapeSectionOutputList<string> GetCells(VASS.Query.SingleSectionQuery query, IVisio.Shape shape, CellValueType type)
        {
            var surface = new SurfaceTarget(shape);
            if (type == CellValueType.Formula)
            {
                return query.GetFormulas(surface);
            }
            else
            {
                return query.GetResults<string>(surface);
            }
        }

        private VASS.Query.ShapesSectionsOutputList<string> GetCells(VASS.Query.SingleSectionQuery query, IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var surface = new SurfaceTarget(page);
            if (type == CellValueType.Formula)
            {
                return query.GetFormulas(surface, shapeids);
            }
            else
            {
                return query.GetResults<string>(surface, shapeids);
            }
        }

        private VASS.Query.Row<string> GetCells(VASS.Query.CellQuery query, IVisio.Shape shape, CellValueType type)
        {
            var surface = new SurfaceTarget(shape);
            if (type == CellValueType.Formula)
            {
                return query.GetFormulas(surface);
            }
            else
            {
                return query.GetResults<string>(surface);
            }
        }

        private VASS.Query.RowList<string> GetCells(VASS.Query.CellQuery query, IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var surface = new SurfaceTarget(page);
            if (type == CellValueType.Formula)
            {
                return query.GetFormulas(surface, shapeids);
            }
            else
            {
                return query.GetResults<string>(surface, shapeids);
            }
        }

    }
}