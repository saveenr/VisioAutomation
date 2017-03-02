using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class XFormCellsReader : SingleRowReader<VisioAutomation.Shapes.XFormCells>
    {
        public CellColumn Width { get; set; }
        public CellColumn Height { get; set; }
        public CellColumn PinX { get; set; }
        public CellColumn PinY { get; set; }
        public CellColumn LocPinX { get; set; }
        public CellColumn LocPinY { get; set; }
        public CellColumn Angle { get; set; }
        
        public XFormCellsReader() 
        {
            this.PinX = this.query.AddCell(SrcConstants.PinX, nameof(SrcConstants.PinX));
            this.PinY = this.query.AddCell(SrcConstants.PinY, nameof(SrcConstants.PinY));
            this.LocPinX = this.query.AddCell(SrcConstants.LocPinX, nameof(SrcConstants.LocPinX));
            this.LocPinY = this.query.AddCell(SrcConstants.LocPinY, nameof(SrcConstants.LocPinY));
            this.Width = this.query.AddCell(SrcConstants.Width, nameof(SrcConstants.Width));
            this.Height = this.query.AddCell(SrcConstants.Height, nameof(SrcConstants.Height));
            this.Angle = this.query.AddCell(SrcConstants.Angle, nameof(SrcConstants.Angle));
        }

        public override XFormCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.XFormCells();
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