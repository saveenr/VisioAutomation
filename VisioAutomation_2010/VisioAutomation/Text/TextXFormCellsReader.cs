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
            this.TxtPinX = this.query.AddCell(SrcConstants.TextXFormPinX, nameof(SrcConstants.TextXFormPinX));
            this.TxtPinY = this.query.AddCell(SrcConstants.TextXFormPinY, nameof(SrcConstants.TextXFormPinY));
            this.TxtLocPinX = this.query.AddCell(SrcConstants.TextXFormLocPinX, nameof(SrcConstants.TextXFormLocPinX));
            this.TxtLocPinY = this.query.AddCell(SrcConstants.TextXFormLocPinY, nameof(SrcConstants.TextXFormLocPinY));
            this.TxtWidth = this.query.AddCell(SrcConstants.TextXFormWidth, nameof(SrcConstants.TextXFormWidth));
            this.TxtHeight = this.query.AddCell(SrcConstants.TextXFormHeight, nameof(SrcConstants.TextXFormHeight));
            this.TxtAngle = this.query.AddCell(SrcConstants.TextXFormAngle, nameof(SrcConstants.TextXFormAngle));

        }

        public override Text.TextXFormCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.TextXFormCells();
            cells.PinX = row[this.TxtPinX];
            cells.PinY = row[this.TxtPinY];
            cells.LocPinX = row[this.TxtLocPinX];
            cells.LocPinY = row[this.TxtLocPinY];
            cells.Width = row[this.TxtWidth];
            cells.Height = row[this.TxtHeight];
            cells.Angle = row[this.TxtAngle];
            return cells;
        }
    }
}