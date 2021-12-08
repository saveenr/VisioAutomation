using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

using ROW= VisioAutomation.ShapeSheet.Data.DataRow<string>;
using ROWS= VisioAutomation.ShapeSheet.Data.DataRows<string>;
using ROWGROUP = VisioAutomation.ShapeSheet.Data.DataRowGroup<string>;
using ROWGROUPS = VisioAutomation.ShapeSheet.Data.DataRowGroups<string>;
using COLS = VisioAutomation.ShapeSheet.Data.DataColumns;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecordBuilder<TREC> where TREC : CellRecord, new()
    {
        public readonly CellRecordQueryType CellRecordQueryType;
        protected Query.CellQuery cellquery;
        protected Query.SectionQuery sectionquery;
        private System.Func<ROW, COLS, TREC> func_row_to_rec;

        private CellRecordBuilder()
        {
            this.cellquery = null;
            this.sectionquery = null;
        }

        protected CellRecordBuilder(CellRecordQueryType cell_record_query_type, System.Func<ROW,COLS, TREC> func_row_to_rec)
        {
            this.func_row_to_rec = func_row_to_rec;

            //VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols

            var temp_cells = new TREC();
            Data.DataColumns cols;

            this.CellRecordQueryType = cell_record_query_type;
            if (cell_record_query_type == CellRecordQueryType.CellQuery)
            {
                this.cellquery = new Query.CellQuery();
                cols = this.cellquery.Columns;
            }
            else if (cell_record_query_type == CellRecordQueryType.SectionQuery)
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

        public CellRecords<TREC> GetCellsMultipleShapesSingleRow(
            IVisio.Page page,
            IList<int> shapeids,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordQueryType.CellQuery);
            var records = new CellRecords<TREC>(shapeids.Count);
            var cols = this.cellquery.Columns;
            ROWS rows = this.__cellquery_multipleshapes(page, shapeids, type);
            foreach (var row in rows)
            {
                var record = this.func_row_to_rec(row, cols);
                records.Add(record);
            }

            return records;
        }

        private void _enforce_category(CellRecordQueryType query_type)
        {
            if (this.CellRecordQueryType != query_type)
            {
                throw new Exceptions.InternalAssertionException();
            }
        }

        public TREC GetCellsSingleShapeSingleRow(
            IVisio.Shape shape,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordQueryType.CellQuery);
            var rows = this.__cellquery_singleshape(shape, type);
            var cols = this.cellquery.Columns;
            var first_row = rows[0];
            var record = this.func_row_to_rec(first_row, cols);
            return record;
        }

        public CellRecordsGroup<TREC> GetCellsMultipleShapesMultipleRows(
            IVisio.Page page,
            Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordQueryType.SectionQuery);
            var sec_cols = this.sectionquery[0];

            var rowgroups =
                __sectionquery_multiplerows(sectionquery, page, shapeidpairs, type);
            var list_cellrecords = new CellRecordsGroup<TREC>(shapeidpairs.Count);
            foreach (var rowgroup in rowgroups)
            {
                var first_rowgroup = rowgroup[0];
                var records = this.__sectionshaperows_to_cellrecords(first_rowgroup, sec_cols);
                list_cellrecords.Add(records);
            }

            return list_cellrecords;
        }

        public CellRecords<TREC> GetCellsSingleShapeMultipleRows(
            IVisio.Shape shape,
            Core.CellValueType type)
        {
            this._enforce_category(CellRecordQueryType.SectionQuery);
            var sec_cols = this.sectionquery[0];
            var rowgroup = __sectionquery_singleshape(shape, type);
            var first_rows = rowgroup[0];
            var records = this.__sectionshaperows_to_cellrecords(first_rows, sec_cols);
            return records;
        }

        private CellRecords<TREC> __sectionshaperows_to_cellrecords(
            ROWS rows,
            COLS cols)
        {
            var records = new CellRecords<TREC>(rows.Count);
            foreach (var section_row in rows)
            {
                var record = this.func_row_to_rec(section_row, cols);
                records.Add(record);
            }

            return records;
        }

        private ROWGROUP __sectionquery_singleshape(
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

        private ROWGROUPS __sectionquery_multiplerows(
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