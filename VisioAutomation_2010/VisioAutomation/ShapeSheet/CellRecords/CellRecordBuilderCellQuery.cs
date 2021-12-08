using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Data;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecordBuilderCellQuery<TREC> where TREC : CellRecord, new()
    {
        protected Query.CellQuery cellquery;
        private System.Func<DataRow<string>, DataColumns, TREC> func_row_to_rec;

        private CellRecordBuilderCellQuery()
        {
            this.cellquery = null;
        }

        protected CellRecordBuilderCellQuery(System.Func<DataRow<string>, DataColumns, TREC> func_row_to_rec)
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
            Microsoft.Office.Interop.Visio.Page page,
            IList<int> shapeids,
            Core.CellValueType type)
        {
            var records = new CellRecords<TREC>(shapeids.Count);
            var cols = this.cellquery.Columns;
            DataRows<string> rows = this.__cellquery_multipleshapes(page, shapeids, type);
            foreach (var row in rows)
            {
                var record = this.func_row_to_rec(row, cols);
                records.Add(record);
            }

            return records;
        }

        public TREC GetCellsSingleShapeSingleRow(
            Microsoft.Office.Interop.Visio.Shape shape,
            Core.CellValueType type)
        {
            var rows = this.__cellquery_singleshape(shape, type);
            var cols = this.cellquery.Columns;
            var first_row = rows[0];
            var record = this.func_row_to_rec(first_row, cols);
            return record;
        }


        private CellRecords<TREC> __sectionshaperows_to_cellrecords(
            DataRows<string> rows,
            DataColumns cols)
        {
            var records = new CellRecords<TREC>(rows.Count);
            foreach (var section_row in rows)
            {
                var record = this.func_row_to_rec(section_row, cols);
                records.Add(record);
            }

            return records;
        }

        private DataRows<string> __cellquery_singleshape(
            Microsoft.Office.Interop.Visio.Shape shape,
            Core.CellValueType type)
        {
            DataRows<string> rows = type switch
            {
                Core.CellValueType.Formula => this.cellquery.GetFormulas(shape),
                Core.CellValueType.Result => this.cellquery.GetResults<string>(shape),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rows;
        }

        private DataRows<string> __cellquery_multipleshapes(
            Microsoft.Office.Interop.Visio.Page page,
            IList<int> shapeids,
            Core.CellValueType type)
        {
            DataRows<string> rows = type switch
            {
                Core.CellValueType.Formula => this.cellquery.GetFormulas(page, shapeids),
                Core.CellValueType.Result => this.cellquery.GetResults<string>(page, shapeids),
                _ => throw new System.ArgumentOutOfRangeException(nameof(type))
            };
            return rows;
        }
    }
}