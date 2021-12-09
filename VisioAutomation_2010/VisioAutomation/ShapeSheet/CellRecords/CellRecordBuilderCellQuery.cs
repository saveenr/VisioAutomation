using System.Collections.Generic;

using IVisio=Microsoft.Office.Interop.Visio;

using ROW = VisioAutomation.ShapeSheet.Data.DataRow<string>;
using ROWS = VisioAutomation.ShapeSheet.Data.DataRows<string>;
using ROWGROUP = VisioAutomation.ShapeSheet.Data.DataRowGroup<string>;
using ROWGROUPS = VisioAutomation.ShapeSheet.Data.DataRowGroups<string>;
using COLS = VisioAutomation.ShapeSheet.Data.DataColumns;

namespace VisioAutomation.ShapeSheet.CellRecords
{
    public abstract class CellRecordBuilderCellQuery<TREC> where TREC : CellRecord, new()
    {
        protected Query.CellQuery cellquery;
        private System.Func<ROW, COLS, TREC> func_row_to_rec;

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