using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class ShapeXFormCellsReader : ReaderSingleRow<VisioAutomation.Shapes.ShapeXFormCells>
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
            this.PinX = this.query.AddColumn(SrcConstants.XFormPinX, nameof(SrcConstants.XFormPinX));
            this.PinY = this.query.AddColumn(SrcConstants.XFormPinY, nameof(SrcConstants.XFormPinY));
            this.LocPinX = this.query.AddColumn(SrcConstants.XFormLocPinX, nameof(SrcConstants.XFormLocPinX));
            this.LocPinY = this.query.AddColumn(SrcConstants.XFormLocPinY, nameof(SrcConstants.XFormLocPinY));
            this.Width = this.query.AddColumn(SrcConstants.XFormWidth, nameof(SrcConstants.XFormWidth));
            this.Height = this.query.AddColumn(SrcConstants.XFormHeight, nameof(SrcConstants.XFormHeight));
            this.Angle = this.query.AddColumn(SrcConstants.XFormAngle, nameof(SrcConstants.XFormAngle));
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