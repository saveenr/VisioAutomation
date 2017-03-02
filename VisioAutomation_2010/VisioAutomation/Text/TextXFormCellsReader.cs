using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    class TextXFormCellsReader : SingleRowReader<Text.TextXFormCells>
    {
        public CellColumn TxtWidth { get; set; }
        public CellColumn TxtHeight { get; set; }
        public CellColumn TxtPinX { get; set; }
        public CellColumn TxtPinY { get; set; }
        public CellColumn TxtLocPinX { get; set; }
        public CellColumn TxtLocPinY { get; set; }
        public CellColumn TxtAngle { get; set; }

        public TextXFormCellsReader()
        {
            this.TxtPinX = this.query.AddCell(SrcConstants.TxtPinX, nameof(SrcConstants.TxtPinX));
            this.TxtPinY = this.query.AddCell(SrcConstants.TxtPinY, nameof(SrcConstants.TxtPinY));
            this.TxtLocPinX = this.query.AddCell(SrcConstants.TxtLocPinX, nameof(SrcConstants.TxtLocPinX));
            this.TxtLocPinY = this.query.AddCell(SrcConstants.TxtLocPinY, nameof(SrcConstants.TxtLocPinY));
            this.TxtWidth = this.query.AddCell(SrcConstants.TxtWidth, nameof(SrcConstants.TxtWidth));
            this.TxtHeight = this.query.AddCell(SrcConstants.TxtHeight, nameof(SrcConstants.TxtHeight));
            this.TxtAngle = this.query.AddCell(SrcConstants.TxtAngle, nameof(SrcConstants.TxtAngle));

        }

        public override Text.TextXFormCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
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