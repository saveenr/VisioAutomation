using VA = VisioAutomation;

namespace VisioAutomation.Effects
{
    public class ThreePointGradientFill
    {
        public VA.Drawing.DirectionRelative Direction { get; set; }
        public VA.Drawing.ColorRGB SideColor { get; set; }
        public VA.Drawing.ColorRGB Corner1Color { get; set; }
        public VA.Drawing.ColorRGB Corner2Color { get; set; }
        public double SideTransparency { get; set; }
        public double Corner1Transparency { get; set; }
        public double Corner2Transparency { get; set; }
        
        public ThreePointGradientFill()
        {
            this.Direction = VA.Drawing.DirectionRelative.Right;
            this.SideColor = new VA.Drawing.ColorRGB(0, 0, 255);
            this.SideTransparency = 0.0;
            this.Corner1Color = new VA.Drawing.ColorRGB(255, 0, 0);
            this.Corner1Transparency = 0.0;
            this.Corner2Color = new VA.Drawing.ColorRGB(0, 255, 0);
            this.Corner2Transparency = 0.0;
        }

        public VA.Format.ShapeFormatCells GetFormat()
        {
            var filldef = this;

            var fgpat = VA.Format.FillPattern.RadialCenter;
            var shadowpat = VA.Format.FillPattern.RadialCenter;

            if (filldef.Direction == VA.Drawing.DirectionRelative.Right)
            {
                fgpat = VA.Format.FillPattern.LinearLeftToRight;
                shadowpat = VA.Format.FillPattern.LinearTopToBottom;
            }
            else if (filldef.Direction == VA.Drawing.DirectionRelative.Down)
            {
                fgpat = VA.Format.FillPattern.LinearTopToBottom;
                shadowpat = VA.Format.FillPattern.LinearRightToLeft;
            }
            else if (filldef.Direction == VA.Drawing.DirectionRelative.Left)
            {
                fgpat = VA.Format.FillPattern.LinearRightToLeft;
                shadowpat = VA.Format.FillPattern.LinearBottomToTop;
            }
            else if (filldef.Direction == VA.Drawing.DirectionRelative.Up)
            {
                fgpat = VA.Format.FillPattern.LinearBottomToTop;
                shadowpat = VA.Format.FillPattern.LinearLeftToRight;
            }

            const double max_transparency = 1.0;

            var format = new VA.Format.ShapeFormatCells();
            format.FillPattern = (int)fgpat;
            format.FillForegnd = VA.Convert.ColorToFormulaRGB(filldef.SideColor);
            format.FillBkgndTrans = filldef.SideTransparency;
            format.FillBkgnd = VA.Convert.ColorToFormulaRGB(filldef.SideColor);
            format.FillBkgndTrans = max_transparency;

            format.ShdwPattern = (int)shadowpat;
            format.ShdwBkgnd = VA.Convert.ColorToFormulaRGB(filldef.Corner2Color);
            format.ShdwBkgndTrans = filldef.Corner2Transparency;
            format.ShdwForegnd = VA.Convert.ColorToFormulaRGB(filldef.Corner1Color);
            format.ShdwForegndTrans = filldef.Corner1Transparency;
            format.ShapeShdwType = 1;
            format.ShapeShdwOffsetX = 0.0;
            format.ShapeShdwOffsetY = 0.0;
            format.ShapeShdwScaleFactor = 1.0;

            return format;
        }
    }
}
