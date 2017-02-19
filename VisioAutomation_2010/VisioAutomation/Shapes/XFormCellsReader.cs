using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Shapes
{
    class XFormCellsReader : SingleRowReader<VisioAutomation.Shapes.XFormCells>
    {
        public ColumnCell Width { get; set; }
        public ColumnCell Height { get; set; }
        public ColumnCell PinX { get; set; }
        public ColumnCell PinY { get; set; }
        public ColumnCell LocPinX { get; set; }
        public ColumnCell LocPinY { get; set; }
        public ColumnCell Angle { get; set; }
        
        public XFormCellsReader() 
        {
            this.PinX = this.query.AddCell(SRCCON.PinX, nameof(SRCCON.PinX));
            this.PinY = this.query.AddCell(SRCCON.PinY, nameof(SRCCON.PinY));
            this.LocPinX = this.query.AddCell(SRCCON.LocPinX, nameof(SRCCON.LocPinX));
            this.LocPinY = this.query.AddCell(SRCCON.LocPinY, nameof(SRCCON.LocPinY));
            this.Width = this.query.AddCell(SRCCON.Width, nameof(SRCCON.Width));
            this.Height = this.query.AddCell(SRCCON.Height, nameof(SRCCON.Height));
            this.Angle = this.query.AddCell(SRCCON.Angle, nameof(SRCCON.Angle));
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