using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using VisioAutomation.ShapeSheet.Queries.Columns;

namespace VisioAutomation.Text
{
    class TextXFormCellsQuery : CellGroupSingleRowQuery<Text.TextXFormCells>
    {
        public ColumnQuery TxtWidth { get; set; }
        public ColumnQuery TxtHeight { get; set; }
        public ColumnQuery TxtPinX { get; set; }
        public ColumnQuery TxtPinY { get; set; }
        public ColumnQuery TxtLocPinX { get; set; }
        public ColumnQuery TxtLocPinY { get; set; }
        public ColumnQuery TxtAngle { get; set; }

        public TextXFormCellsQuery()
        {
            this.TxtPinX = this.query.AddCell(SRCConstants.TxtPinX, nameof(SRCConstants.TxtPinX));
            this.TxtPinY = this.query.AddCell(SRCConstants.TxtPinY, nameof(SRCConstants.TxtPinY));
            this.TxtLocPinX = this.query.AddCell(SRCConstants.TxtLocPinX, nameof(SRCConstants.TxtLocPinX));
            this.TxtLocPinY = this.query.AddCell(SRCConstants.TxtLocPinY, nameof(SRCConstants.TxtLocPinY));
            this.TxtWidth = this.query.AddCell(SRCConstants.TxtWidth, nameof(SRCConstants.TxtWidth));
            this.TxtHeight = this.query.AddCell(SRCConstants.TxtHeight, nameof(SRCConstants.TxtHeight));
            this.TxtAngle = this.query.AddCell(SRCConstants.TxtAngle, nameof(SRCConstants.TxtAngle));

        }

        public override Text.TextXFormCells CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new Text.TextXFormCells();
            cells.TxtPinX = row[this.TxtPinX];
            cells.TxtPinY = row[this.TxtPinY];
            cells.TxtLocPinX = row[this.TxtLocPinX];
            cells.TxtLocPinY = row[this.TxtLocPinY];
            cells.TxtWidth = row[this.TxtWidth];
            cells.TxtHeight = row[this.TxtHeight];
            cells.TxtAngle = row[this.TxtAngle];
            return cells;
        }
    }
}