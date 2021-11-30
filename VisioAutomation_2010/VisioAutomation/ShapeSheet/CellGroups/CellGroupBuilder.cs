using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupBuilder<TGroup> where TGroup : CellGroup, new()
    {
        public readonly CellGroupBuilderType Type;
        protected Query.CellQuery query_cells_singlerow;
        protected Query.SectionQuery query_sections_multirow;

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
                cols = this.query_sections_multirow.Add(temp_cells.GetCellMetadata().First().Src);
            }
            else
            {
                throw new Exceptions.InternalAssertionException();
            }

            foreach (var pair in temp_cells.GetCellMetadata())
            {
                cols.Add(pair.Src, pair.Name);
            }

        }

        public abstract TGroup ToCellGroup(Query.Row<string> row, Query.Columns cols);

        public List<TGroup> GetCellsSingleRow(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            this._enforce_type(CellGroupBuilderType.SingleRow);
            var cellgroups = new List<TGroup>(shapeids.Count);
            var cols = this.query_cells_singlerow.Columns;
            var rows_for_shapes = this.__GetCells(query_cells_singlerow, page, shapeids, type);
            foreach (var row in rows_for_shapes)
            {
                var cellgroup = this.ToCellGroup(row, cols);
                cellgroups.Add(cellgroup);
            }
            return cellgroups;
        }

        private void _enforce_type(CellGroupBuilderType t)
        {
            if (this.Type != t)
            {
                throw new Exceptions.InternalAssertionException();
            }
        }

        public TGroup GetCellsSingleRow(IVisio.Shape shape, Core.CellValueType type)
        {
            this._enforce_type(CellGroupBuilderType.SingleRow);
            var cellqueryresult = this.__GetCells(query_cells_singlerow, shape, type);
            var cols = this.query_cells_singlerow.Columns;
            var first_row = cellqueryresult[0];
            var cells = this.ToCellGroup(first_row, cols);
            return cells;
        }
        
        public List<List<TGroup>> GetCellsMultiRow(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, Core.CellValueType type)
        {
            this._enforce_type(CellGroupBuilderType.MultiRow);
            var sec_cols = this.query_sections_multirow[0];

            var cellqueryresult = __GetCells(query_sections_multirow,page, shapeidpairs, type);
            var list_cellgroups = new List<List<TGroup>>(shapeidpairs.Count);
            foreach (var data_for_shape in cellqueryresult)
            {
                var first_section_results = data_for_shape[0];
                var cellgroups = this._sectionshaperows_to_cellgroups(first_section_results,sec_cols);
                list_cellgroups.Add(cellgroups);
            }
            return list_cellgroups;
        }

        public List<TGroup> GetCellsMultiRow(IVisio.Shape shape, Core.CellValueType type)
        {
            this._enforce_type(CellGroupBuilderType.MultiRow);
            var sec_cols = this.query_sections_multirow[0];
            var cellqueryresult = __GetCells(query_sections_multirow, shape, type);
            var shape0_sectionshaperows0 = cellqueryresult[0];
            var cellgroups = this._sectionshaperows_to_cellgroups(shape0_sectionshaperows0,sec_cols);
            return cellgroups;
        }

        private List<TGroup> _sectionshaperows_to_cellgroups(Query.SectionShapeRows<string> section_rows, Query.Columns cols)
        {
            var cellgroups = new List<TGroup>(section_rows.Count);
            foreach (var section_row in section_rows)
            {
                var cellgroup = this.ToCellGroup(section_row,cols);
                cellgroups.Add(cellgroup);
            }
            return cellgroups;
        }

        private Query.SectionQueryShapeResults<string> __GetCells(Query.SectionQuery query, IVisio.Shape shape, Core.CellValueType type)
        {
            var results = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(shape),
                Core.CellValueType.Result => query.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return results;
        }

        private Query.SectionQueryResults<string> __GetCells(Query.SectionQuery query, IVisio.Page page, Core.ShapeIDPairs shapeidpairs, Core.CellValueType type)
        {
            var results = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(page, shapeidpairs),
                Core.CellValueType.Result => query.GetResults<string>(page, shapeidpairs),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return results;
        }

        private Query.CellQueryResults<string> __GetCells(Query.CellQuery query, IVisio.Shape shape, Core.CellValueType type)
        {
            var results = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(shape),
                Core.CellValueType.Result => query.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return results;
        }

        private Query.CellQueryResults<string> __GetCells(Query.CellQuery query, IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var results = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(page, shapeids),
                Core.CellValueType.Result => query.GetResults<string>(page, shapeids),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return results;
        }
    }
}