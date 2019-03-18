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
        protected VASS.Query.SectionQuery query_sections_multirow;

        private CellGroupBuilder()
        {
            this.query_cells_singlerow = null;
            this.query_sections_multirow = null;
        }

        protected CellGroupBuilder(CellGroupBuilderType type)
        {
            var temp_cells = new TGroup();
            Query.Columns cols;

            this.Type = type;
            if (type == CellGroupBuilderType.SingleRow)
            {
                this.query_cells_singlerow = new Query.CellQuery();
                cols = this.query_cells_singlerow.Columns;
            }
            else if (type == CellGroupBuilderType.MultiRow)
            {
                this.query_sections_multirow = new Query.SectionQuery();
                cols = this.query_sections_multirow.Add(temp_cells.CellMetadata.First().Src);
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

        public abstract TGroup ToCellGroup(VisioAutomation.ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols);

        public List<TGroup> GetCellsSingleRow(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.SingleRow);
            var cellgroups = new List<TGroup>(shapeids.Count);
            var cols = this.query_cells_singlerow.Columns;
            var rows_for_shapes = this.GetCells(query_cells_singlerow, page, shapeids, type);
            foreach (var row in rows_for_shapes)
            {
                var cellgroup = this.ToCellGroup(row, cols);
                cellgroups.Add(cellgroup);
            }
            return cellgroups;
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
            var cellqueryresult = this.GetCells(query_cells_singlerow, shape, type);
            var cols = this.query_cells_singlerow.Columns;
            var first_row = cellqueryresult[0];
            var cells = this.ToCellGroup(first_row, cols);
            return cells;
        }
        
        public List<List<TGroup>> GetCellsMultiRow(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.MultiRow);
            var sec_cols = this.query_sections_multirow[0];

            var cellqueryresult = GetCells(query_sections_multirow,page, shapeids, type);
            var list_cellgroups = new List<List<TGroup>>(shapeids.Count);
            foreach (var data_for_shape in cellqueryresult)
            {
                var first_section_results = data_for_shape[0];
                var cellgroups = this._section_result_to_cell_groups(first_section_results,sec_cols);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }

        public List<TGroup> GetCellsMultiRow(IVisio.Shape shape, CellValueType type)
        {
            this.EnforceType(CellGroupBuilderType.MultiRow);
            var sec_cols = this.query_sections_multirow[0];
            var cellqueryresult = GetCells(query_sections_multirow, shape, type);
            var first_section = cellqueryresult[0];
            var cellgroups = this._section_result_to_cell_groups(first_section,sec_cols);
            return cellgroups;
        }

        private List<TGroup> _section_result_to_cell_groups(VASS.Query.ShapeSectionRows<string> section_rows, VisioAutomation.ShapeSheet.Query.Columns cols)
        {
            var cellgroups = new List<TGroup>(section_rows.Count);
            foreach (var section_row in section_rows)
            {
                var cellgroup = this.ToCellGroup(section_row,cols);
                cellgroups.Add(cellgroup);
            }
            return cellgroups;
        }

        private VASS.Query.SectionQueryShapeResults<string> GetCells(VASS.Query.SectionQuery query, IVisio.Shape shape, CellValueType type)
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

        private VASS.Query.SectionQueryResults<string> GetCells(VASS.Query.SectionQuery query, IVisio.Page page, IList<int> shapeids, CellValueType type)
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

        private VASS.Query.CellQueryResults<string> GetCells(VASS.Query.CellQuery query, IVisio.Shape shape, CellValueType type)
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

        private VASS.Query.CellQueryResults<string> GetCells(VASS.Query.CellQuery query, IVisio.Page page, IList<int> shapeids, CellValueType type)
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