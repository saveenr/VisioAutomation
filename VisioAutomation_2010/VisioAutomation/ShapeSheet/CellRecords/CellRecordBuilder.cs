using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

using ROWS= VisioAutomation.ShapeSheet.Data.DataRows<string>;
using COLS = VisioAutomation.ShapeSheet.Data.DataColumns;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecordBuilder<TRecord> where TRecord : CellRecord, new()
    {
        public readonly CellRecordCategory Type;
        protected Query.CellQuery cellquery;
        protected Query.SectionQuery sectionquery;

        private CellRecordBuilder()
        {
            this.cellquery = null;
            this.sectionquery = null;
        }

        protected CellRecordBuilder(CellRecordCategory type)
        {
            var temp_cells = new TRecord();
            Data.DataColumns cols;

            this.Type = type;
            if (type == CellRecordCategory.SingleRow)
            {
                this.cellquery = new Query.CellQuery();
                cols = this.cellquery.Columns;
            }
            else if (type == CellRecordCategory.MultiRow)
            {
                this.sectionquery = new Query.SectionQuery();
                cols = this.sectionquery.Add(temp_cells.GetCellMetadata().First().Src);
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

        public abstract TRecord ToCellRecord(Data.DataRow<string> row, Data.DataColumns cols);

        public List<TRecord> GetCellsMultipleShapesSingleRow(
            IVisio.Page page, 
            IList<int> shapeids,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.SingleRow);
            var records = new List<TRecord>(shapeids.Count);
            var cols = this.cellquery.Columns;
            ROWS rows = this.__cellquery_multipleshapes(page, shapeids, type);
            foreach (var row in rows)
            {
                var record = this.ToCellRecord(row, cols);
                records.Add(record);
            }

            return records;
        }

        private void _enforce_category(CellRecordCategory category)
        {
            if (this.Type != category)
            {
                throw new Exceptions.InternalAssertionException();
            }
        }

        public TRecord GetCellsSingleShapeSingleRow(
            IVisio.Shape shape,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.SingleRow);
            var rows = this.__cellquery_singleshape(shape, type);
            var cols = this.cellquery.Columns;
            var first_row = rows[0];
            var record = this.ToCellRecord(first_row, cols);
            return record;
        }

        public List<List<TRecord>> GetCellsMultipleShapesMultipleRows(
            IVisio.Page page, 
            Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.MultiRow);
            var sec_cols = this.sectionquery[0];

            var rowgroups =
                __sectionquery_multiplerows(sectionquery, page, shapeidpairs, type);
            var list_cellrecords = new List<List<TRecord>>(shapeidpairs.Count);
            foreach (var rowgroup in rowgroups)
            {
                var first_rowgroup = rowgroup[0];
                var records = this.__sectionshaperows_to_cellrecords(first_rowgroup, sec_cols);
                list_cellrecords.Add(records);
            }

            return list_cellrecords;
        }

        public List<TRecord> GetCellsSingleShapeMultipleRows(
            IVisio.Shape shape, 
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordCategory.MultiRow);
            var sec_cols = this.sectionquery[0];
            var rowgroup = __sectionquery_singleshape(shape, type);
            var first_rows = rowgroup[0];
            var records = this.__sectionshaperows_to_cellrecords(first_rows, sec_cols);
            return records;
        }

        private List<TRecord> __sectionshaperows_to_cellrecords(
            ROWS rows,
            COLS cols)
        {
            var records = new List<TRecord>(rows.Count);
            foreach (var section_row in rows)
            {
                var record = this.ToCellRecord(section_row, cols);
                records.Add(record);
            }

            return records;
        }

        private Data.DataRowGroup<string> __sectionquery_singleshape(
            IVisio.Shape shape, 
            Core.CellValueType type)
        {
            var rowgroup = type switch
            {
                Core.CellValueType.Formula => this.sectionquery.GetFormulas(shape),
                Core.CellValueType.Result => this.sectionquery.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rowgroup;
        }

        private Data.DataRowGroups<string> __sectionquery_multiplerows(
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

        private ROWS __cellquery_singleshape(
            IVisio.Shape shape, 
            Core.CellValueType type)
        {
            ROWS rows = type switch
            {
                Core.CellValueType.Formula => this.cellquery.GetFormulas(shape),
                Core.CellValueType.Result => this.cellquery.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rows;
        }

        private ROWS __cellquery_multipleshapes(
            IVisio.Page page, 
            IList<int> shapeids, 
            Core.CellValueType type)
        {
            ROWS rows = type switch
            {
                Core.CellValueType.Formula => this.cellquery.GetFormulas(page, shapeids),
                Core.CellValueType.Result => this.cellquery.GetResults<string>(page, shapeids),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rows;
        }
    }
}