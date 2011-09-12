using System;
using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public static class MathUtil
    {
        public static VA.Drawing.Size Max(VA.Drawing.Size a, VA.Drawing.Size b)
        {
            return new VA.Drawing.Size(Math.Max(a.Width, b.Width),
                            Math.Max(a.Height, b.Height));
        }

        public static VA.Drawing.Size SnapToNearestValue(VA.Drawing.Size size, VA.Drawing.Size snapsize)
        {
            return new VA.Drawing.Size(VA.Internal.MathUtil.Round(size.Width, snapsize.Width),
                            VA.Internal.MathUtil.Round(size.Height, snapsize.Height));
        }

        public static VA.Drawing.Point Round(VA.Drawing.Point p, double xd, double yd)
        {
            return new Drawing.Point(VA.Internal.MathUtil.Round(p.X, xd),
                             VA.Internal.MathUtil.Round(p.Y, yd));
        }
    }
}