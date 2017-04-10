using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class ShapeXFormCellsReader : SingleRowReader<VisioAutomation.Shapes.ShapeXFormCells>
    {
        public CellColumn Width { get; set; }
        public CellColumn Height { get; set; }
        public CellColumn PinX { get; set; }
        public CellColumn PinY { get; set; }
        public CellColumn LocPinX { get; set; }
        public CellColumn LocPinY { get; set; }
        public CellColumn Angle { get; set; }
        
        public ShapeXFormCellsReader() 
        {
            this.PinX = this.query.AddCell(SrcConstants.XFormPinX, nameof(SrcConstants.XFormPinX));
            this.PinY = this.query.AddCell(SrcConstants.XFormPinY, nameof(SrcConstants.XFormPinY));
            this.LocPinX = this.query.AddCell(SrcConstants.XFormLocPinX, nameof(SrcConstants.XFormLocPinX));
            this.LocPinY = this.query.AddCell(SrcConstants.XFormLocPinY, nameof(SrcConstants.XFormLocPinY));
            this.Width = this.query.AddCell(SrcConstants.XFormWidth, nameof(SrcConstants.XFormWidth));
            this.Height = this.query.AddCell(SrcConstants.XFormHeight, nameof(SrcConstants.XFormHeight));
            this.Angle = this.query.AddCell(SrcConstants.XFormAngle, nameof(SrcConstants.XFormAngle));
        }

        public override ShapeXFormCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.ShapeXFormCells();
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