using System;
using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public class SnappingGrid
    {
        public VA.Drawing.Size SnapSize { get; private set; }
        
        public SnappingGrid(double w, double h)
        {
            this.SnapSize = new VA.Drawing.Size(w, h);
        }

        public SnappingGrid( VA.Drawing.Size size)
        {
            this.SnapSize = size;
        }

        public VA.Drawing.Size Snap(VA.Drawing.Size size)
        {
            var w = VA.Internal.MathUtil.Round(size.Width, this.SnapSize.Width);
            var h = VA.Internal.MathUtil.Round(size.Height, this.SnapSize.Height);
            return new VA.Drawing.Size(w,h);
        }
    }
}