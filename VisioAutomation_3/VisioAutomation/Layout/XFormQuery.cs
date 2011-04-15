using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout
{
    class XFormQuery : VA.ShapeSheet.Query.CellQuery
    {
        public VA.ShapeSheet.Query.CellQueryColumn Width { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn Height { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn PinX { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn PinY { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LocPinX { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LocPinY { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn Angle { get; set; }

        public XFormQuery() :
            base()
        {
            PinX = this.AddColumn(VA.ShapeSheet.SRCConstants.PinX, "PinX");
            PinY = this.AddColumn(VA.ShapeSheet.SRCConstants.PinY, "PinY");
            LocPinX = this.AddColumn(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
            LocPinY = this.AddColumn(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
            Width = this.AddColumn(VA.ShapeSheet.SRCConstants.Width, "Width");
            Height = this.AddColumn(VA.ShapeSheet.SRCConstants.Height, "Height");
            Angle = this.AddColumn(VA.ShapeSheet.SRCConstants.Angle, "Angle");
        }
    }

}