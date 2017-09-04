using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    class TextXFormCellsReader : ReaderSingleRow<Text.TextXFormCells>
    {
        public CellColumn Width { get; set; }
        public CellColumn Height { get; set; }
        public CellColumn PinX { get; set; }
        public CellColumn PinY { get; set; }
        public CellColumn LocPinX { get; set; }
        public CellColumn LocPinY { get; set; }
        public CellColumn Angle { get; set; }

        public TextXFormCellsReader()
        {
            this.PinX = this.query.AddCell(SrcConstants.TextXFormPinX, nameof(SrcConstants.TextXFormPinX));
            this.PinY = this.query.AddCell(SrcConstants.TextXFormPinY, nameof(SrcConstants.TextXFormPinY));
            this.LocPinX = this.query.AddCell(SrcConstants.TextXFormLocPinX, nameof(SrcConstants.TextXFormLocPinX));
            this.LocPinY = this.query.AddCell(SrcConstants.TextXFormLocPinY, nameof(SrcConstants.TextXFormLocPinY));
            this.Width = this.query.AddCell(SrcConstants.TextXFormWidth, nameof(SrcConstants.TextXFormWidth));
            this.Height = this.query.AddCell(SrcConstants.TextXFormHeight, nameof(SrcConstants.TextXFormHeight));
            this.Angle = this.query.AddCell(SrcConstants.TextXFormAngle, nameof(SrcConstants.TextXFormAngle));

        }

        public override Text.TextXFormCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.TextXFormCells();
            cells.PinX = row[this.PinX];
            cells.PinY = row[this.PinY];
            cells.LocPinX = row[this.LocPinX];
            cells.LocPinY = row[this.LocPinY];
            cells.Width = row[this.Width];
            cells.Height = row[this.Height];
            cells.Angle = row[this.Angle];
            return cells;
        }
    }
}