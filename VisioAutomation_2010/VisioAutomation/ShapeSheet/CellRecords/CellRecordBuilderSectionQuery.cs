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

    public abstract class CellRecordBuilderCellQuery<TREC> where TREC : CellRecord, new()
    {
        protected Query.CellQuery cellquery;
        private System.Func<ROW, COLS, TREC> func_row_to_rec;

        private CellRecordBuilderCellQuery()
        {
            this.cellquery = null;
        }

        protected CellRecordBuilderCellQuery(System.Func<ROW, COLS, TREC> func_row_to_rec)
        {
            this.func_row_to_rec = func_row_to_rec;
            this.cellquery = new Query.CellQuery();
            var querycols = this.cellquery.Columns;
            var temp_cells = new TREC();
            foreach (var item in temp_cells.GetCellMetadata())
            {
                querycols.Add(item.Src, item.Name);
            }
        }

        public CellRecords<TREC> GetCellsMultipleShapesSingleRow(
            IVisio.Page page,
            IList<int> shapeids,
            Core.CellValueType type)
        {
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

        public TREC GetCellsSingleShapeSingleRow(
            IVisio.Shape shape,
            Core.CellValueType type)
        {
            var rows = this.__cellquery_singleshape(shape, type);
            var cols = this.cellquery.Columns;
            var first_row = rows[0];
            var record = this.func_row_to_rec(first_row, cols);
            return record;
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


    public abstract class CellRecordBuilderSectionQuery<TREC> where TREC : CellRecord, new()
    {
        protected Query.SectionQuery sectionquery;
        private System.Func<ROW, COLS, TREC> func_row_to_rec;

        private CellRecordBuilderSectionQuery()
        {
            this.sectionquery = null;
        }

        protected CellRecordBuilderSectionQuery(System.Func<ROW,COLS, TREC> func_row_to_rec)
        {
            this.func_row_to_rec = func_row_to_rec;

            var temp_cells = new TREC();
            Data.DataColumns querycols;
            this.sectionquery = new Query.SectionQuery();
            querycols = this.sectionquery.Add(temp_cells.GetCellMetadata().First().Src);

            foreach (var item in temp_cells.GetCellMetadata())
            {
                querycols.Add(item.Src, item.Name);
            }
        }

        public CellRecordsGroup<TREC> GetCellsMultipleShapesMultipleRows(
            IVisio.Page page,
            Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {

            var sec_cols = this.sectionquery[0];

            var rowgroups =
                __sectionquery_multiplerows(sectionquery, page, shapeidpairs, type);
            var recordgroup = new CellRecordsGroup<TREC>(shapeidpairs.Count);
            foreach (var rowgroup in rowgroups)
            {
                var first_rowgroup = rowgroup[0];
                var records = this.__sectionshaperows_to_cellrecords(first_rowgroup, sec_cols);
                recordgroup.Add(records);
            }

            return recordgroup;
        }

        public CellRecords<TREC> GetCellsSingleShapeMultipleRows(
            IVisio.Shape shape,
            Core.CellValueType type)
        {

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
    }

}