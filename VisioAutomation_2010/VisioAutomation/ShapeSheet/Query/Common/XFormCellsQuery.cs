using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class XFormCellsQuery : CellQuery
    {
        public Query.CellColumn Width { get; set; }
        public Query.CellColumn Height { get; set; }
        public Query.CellColumn PinX { get; set; }
        public Query.CellColumn PinY { get; set; }
        public Query.CellColumn LocPinX { get; set; }
        public Query.CellColumn LocPinY { get; set; }
        public Query.CellColumn Angle { get; set; }

        public XFormCellsQuery()
        {
            this.PinX = this.AddCell(SRCCON.PinX, nameof(SRCCON.PinX));
            this.PinY = this.AddCell(SRCCON.PinY, nameof(SRCCON.PinY));
            this.LocPinX = this.AddCell(SRCCON.LocPinX, nameof(SRCCON.LocPinX));
            this.LocPinY = this.AddCell(SRCCON.LocPinY, nameof(SRCCON.LocPinY));
            this.Width = this.AddCell(SRCCON.Width, nameof(SRCCON.Width));
            this.Height = this.AddCell(SRCCON.Height, nameof(SRCCON.Height));
            this.Angle = this.AddCell(SRCCON.Angle, nameof(SRCCON.Angle));
        }

        public VisioAutomation.Shapes.XFormCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.XFormCells
            {
                PinX = row[this.PinX],
                PinY = row[this.PinY],
                LocPinX = row[this.LocPinX],
                LocPinY = row[this.LocPinY],
                Width = row[this.Width],
                Height = row[this.Height],
                Angle = row[this.Angle]
            };
            return cells;
        }
    }
}