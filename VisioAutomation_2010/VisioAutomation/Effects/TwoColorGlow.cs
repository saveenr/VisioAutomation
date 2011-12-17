using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;

namespace VisioAutomation.Effects
{
    public class TwoColorGlow
    {
        public VA.Drawing.ColorRGB TopColor { get; set; }
        public VA.Drawing.ColorRGB BottomColor { get; set; }
        public double TopTransparency { get; set; }
        public double BottomTransparency { get; set; }
        public double Scale { get; set; }

        public TwoColorGlow()
        {
            this.TopColor = new VA.Drawing.ColorRGB(255, 0, 0);
            this.TopTransparency = 0.0;
            this.BottomColor = new VA.Drawing.ColorRGB(255, 255, 0);
            this.BottomTransparency = 0.0;
            this.Scale = 2.0;
        }

        public VA.Format.ShapeFormatCells GetFormat()
        {
            var fgpat = VA.Format.FillPattern.RadialCenter;
            var shadowpat = VA.Format.FillPattern.RadialCenter;
            const double max_transparency = 1.0;

            var format = new VA.Format.ShapeFormatCells();
            format.FillPattern = (int)fgpat;
            format.FillForegnd = VA.Convert.ColorToFormulaRGB(this.TopColor);
            format.FillBkgndTrans = this.TopTransparency;
            format.FillBkgnd = VA.Convert.ColorToFormulaRGB(this.TopColor);
            format.FillBkgndTrans = max_transparency;
            format.ShdwPattern = (int)shadowpat;
            format.ShdwBkgnd = VA.Convert.ColorToFormulaRGB(this.BottomColor);
            format.ShdwBkgndTrans = max_transparency;
            format.ShdwForegnd = VA.Convert.ColorToFormulaRGB(this.BottomColor);
            format.ShdwForegndTrans = this.BottomTransparency;
            format.ShapeShdwType = 1;
            format.ShapeShdwOffsetX = 0.0;
            format.ShapeShdwOffsetY = 0.0;
            format.ShapeShdwScaleFactor = this.Scale;
            return format;
        }
    }
}
