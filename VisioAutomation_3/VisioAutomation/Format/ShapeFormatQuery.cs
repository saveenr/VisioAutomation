using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    class ShapeFormatQuery : VA.ShapeSheet.Query.CellQuery
    {
        public VA.ShapeSheet.Query.CellQueryColumn FillBkgnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn FillBkgndTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn FillForegnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn FillForegndTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn FillPattern { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwObliqueAngle { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwOffsetX { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwOffsetY { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwScaleFactor { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwType { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShdwBkgnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShdwBkgndTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShdwForegnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShdwForegndTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShdwPattern { get; set; }

        public VA.ShapeSheet.Query.CellQueryColumn BeginArrow { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn BeginArrowSize { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn EndArrow { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn EndArrowSize { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LineColor { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LineCap { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LineColorTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LinePattern { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn LineWeight { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn Rounding { get; set; }
        
        public VA.ShapeSheet.Query.CellQueryColumn CharColor { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn CharColorTrans { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn CharSize { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn CharFont{ get; set; }

        public VA.ShapeSheet.Query.CellQueryColumn TextBkgnd { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn TextBkgndTrans { get; set; }
        
        public ShapeFormatQuery() :
            base()
        {
            this.FillBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgnd, "FillBkgnd");
            this.FillBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgndTrans, "FillBkgndTrans");
            this.FillForegnd = this.AddColumn(VA.ShapeSheet.SRCConstants.FillForegnd, "FillForegnd");
            this.FillForegndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.FillForegndTrans, "FillForegndTrans");
            this.FillPattern = this.AddColumn(VA.ShapeSheet.SRCConstants.FillPattern, "FillPattern");
            this.ShapeShdwObliqueAngle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, "ShapeShdwObliqueAngle");
            this.ShapeShdwOffsetX = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX, "ShapeShdwOffsetX");
            this.ShapeShdwOffsetY = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY, "ShapeShdwOffsetY");
            this.ShapeShdwScaleFactor = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor, "ShapeShdwScaleFactor");
            this.ShapeShdwType = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwType, "ShapeShdwType");
            this.ShdwBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgnd, "ShdwBkgnd ");
            this.ShdwBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShdwBkgndTrans");
            this.ShdwForegnd = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegnd, "ShdwForegnd ");
            this.ShdwForegndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegndTrans, "ShdwForegndTrans");
            this.ShdwPattern = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwPattern, "ShdwPattern");
        
            this.BeginArrow = this.AddColumn(VA.ShapeSheet.SRCConstants.BeginArrow, "BeginArrow");
            this.BeginArrowSize = this.AddColumn(VA.ShapeSheet.SRCConstants.BeginArrowSize, "BeginArrowSize");
            this.EndArrow = this.AddColumn(VA.ShapeSheet.SRCConstants.EndArrow, "EndArrow ");
            this.EndArrowSize = this.AddColumn(VA.ShapeSheet.SRCConstants.EndArrowSize, "EndArrowSize");
            this.LineColor = this.AddColumn(VA.ShapeSheet.SRCConstants.LineColor, "LineColor");
            this.LineCap = this.AddColumn(VA.ShapeSheet.SRCConstants.LineCap, "LineCap");
            this.LineColorTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.LineColorTrans, "LineColorTrans");
            this.LinePattern = this.AddColumn(VA.ShapeSheet.SRCConstants.LinePattern, "LinePattern");
            this.LineWeight = this.AddColumn(VA.ShapeSheet.SRCConstants.LineWeight, "LineWeight");
            this.Rounding = this.AddColumn(VA.ShapeSheet.SRCConstants.Rounding, "Rounding");
            
            this.CharColor = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Color, "CharColor");
            this.CharColorTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_ColorTrans, "CharColorTrans");
            this.CharSize = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size, "CharSize");
            this.CharFont= this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Font, "CharFont");
            
            this.TextBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
            this.TextBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
        }
    }
}