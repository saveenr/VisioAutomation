using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

using ROW= VisioAutomation.ShapeSheet.Data.DataRow<string>;
using ROWS= VisioAutomation.ShapeSheet.Data.DataRows<string>;
using ROWGROUP = VisioAutomation.ShapeSheet.Data.DataRowGroup<string>;
using ROWGROUPS = VisioAutomation.ShapeSheet.Data.DataRowGroups<string>;
using COLS = VisioAutomation.ShapeSheet.Data.DataColumns;

namespace VisioAutomation.ShapeSheet.CellRecords
{
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
                var records = this.rows_to_records(first_rowgroup, sec_cols);
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
            var records = this.rows_to_records(first_rows, sec_cols);
            return records;
        }

        private CellRecords<TREC> rows_to_records(
            ROWS rows,
            COLS cols)
        {
            var records = new CellRecords<TREC>(rows.Count);
            records.AddRange( rows.Select( row => this.func_row_to_rec(row, cols)));
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