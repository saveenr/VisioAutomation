using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

using ROWS= VisioAutomation.ShapeSheet.Data.DataRows<string>;
using COLS = VisioAutomation.ShapeSheet.Data.DataColumns;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecordBuilder<TGroup> where TGroup : CellRecord, new()
    {
        public readonly CellRecordCategory Type;
        protected Query.CellQuery query_cells_singlerow;
        protected Query.SectionQuery query_sections_multirow;

        private CellRecordBuilder()
        {
            this.query_cells_singlerow = null;
            this.query_sections_multirow = null;
        }

        protected CellRecordBuilder(CellRecordCategory type)
        {
            var temp_cells = new TGroup();
            Data.DataColumns cols;

            this.Type = type;
            if (type == CellRecordCategory.SingleRow)
            {
                this.query_cells_singlerow = new Query.CellQuery();
                cols = this.query_cells_singlerow.Columns;
            }
            else if (type == CellRecordCategory.MultiRow)
            {
                this.query_sections_multirow = new Query.SectionQuery();
                cols = this.query_sections_multirow.Add(temp_cells.GetCellMetadata().First().Src);
            }
            else
            {
                throw new Exceptions.InternalAssertionException();
            }

            foreach (var item in temp_cells.GetCellMetadata())
            {
                cols.Add(item.Src, item.Name);
            }
        }

        public abstract TGroup ToCellRecord(Data.DataRow<string> row, Data.DataColumns cols);

        public List<TGroup> GetCellsMultipleShapesSingleRow(
            IVisio.Page page, 
            IList<int> shapeids,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.SingleRow);
            var cell_records = new List<TGroup>(shapeids.Count);
            var cols = this.query_cells_singlerow.Columns;
            ROWS rows =
                this.__QueryCells_MultipleShapes_SingleRow(query_cells_singlerow, page, shapeids, type);
            foreach (var row in rows)
            {
                var cell_record = this.ToCellRecord(row, cols);
                cell_records.Add(cell_record);
            }

            return cell_records;
        }

        private void _enforce_category(CellRecordCategory category)
        {
            if (this.Type != category)
            {
                throw new Exceptions.InternalAssertionException();
            }
        }

        public TGroup GetCellsSingleShapeSingleRow(
            IVisio.Shape shape,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.SingleRow);
            var rows = this.__QueryCells_SingleShape_SingleRow(query_cells_singlerow, shape, type);
            var cols = this.query_cells_singlerow.Columns;
            var first_row = rows[0];
            var cells = this.ToCellRecord(first_row, cols);
            return cells;
        }

        public List<List<TGroup>> GetCellsMultipleShapesMultipleRows(
            IVisio.Page page, 
            Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.MultiRow);
            var sec_cols = this.query_sections_multirow[0];

            var rowgroups =
                __QueryCells_MultipleShapes_MultipleRows(query_sections_multirow, page, shapeidpairs, type);
            var list_cellrecords = new List<List<TGroup>>(shapeidpairs.Count);
            foreach (var rowgroup in rowgroups)
            {
                var first_rowgroup = rowgroup[0];
                var records = this._sectionshaperows_to_cellrecords(first_rowgroup, sec_cols);
                list_cellrecords.Add(records);
            }

            return list_cellrecords;
        }

        public List<TGroup> GetCellsSingleShapeMultipleRows(
            IVisio.Shape shape, 
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.MultiRow);
            var sec_cols = this.query_sections_multirow[0];
            var rowgroup = __QueryCells_SingleShape_MultipleRows(query_sections_multirow, shape, type);
            var first_rows = rowgroup[0];
            var records = this._sectionshaperows_to_cellrecords(first_rows, sec_cols);
            return records;
        }

        private List<TGroup> _sectionshaperows_to_cellrecords(
            ROWS group_cell_value_data_row_collection,
            COLS cols)
        {
            var records = new List<TGroup>(group_cell_value_data_row_collection.Count);
            foreach (var section_row in group_cell_value_data_row_collection)
            {
                var record = this.ToCellRecord(section_row, cols);
                records.Add(record);
            }

            return records;
        }

        private Data.DataRowGroup<string> __QueryCells_SingleShape_MultipleRows(
            Query.SectionQuery query,
            IVisio.Shape shape, 
            Core.CellValueType type)
        {
            var rowgroup = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(shape),
                Core.CellValueType.Result => query.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rowgroup;
        }

        private Data.DataRowGroups<string> __QueryCells_MultipleShapes_MultipleRows(
            Query.SectionQuery query,
            IVisio.Page page, 
            Core.ShapeIDPairs shapeidpairs, 
            Core.CellValueType type)
        {
            var rowgroups = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(page, shapeidpairs),
                Core.CellValueType.Result => query.GetResults<string>(page, shapeidpairs),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rowgroups;
        }

        private ROWS __QueryCells_SingleShape_SingleRow(
            Query.CellQuery query,
            IVisio.Shape shape, 
            Core.CellValueType type)
        {
            ROWS rows = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(shape),
                Core.CellValueType.Result => query.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rows;
        }

        private ROWS __QueryCells_MultipleShapes_SingleRow(
            Query.CellQuery query,
            IVisio.Page page, 
            IList<int> shapeids, 
            Core.CellValueType type)
        {
            ROWS rows = type switch
            {
                Core.CellValueType.Formula => query.GetFormulas(page, shapeids),
                Core.CellValueType.Result => query.GetResults<string>(page, shapeids),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rows;
        }
    }
}