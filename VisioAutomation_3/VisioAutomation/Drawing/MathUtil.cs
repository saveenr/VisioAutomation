using System;
using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public static class MathUtil
    {
        public static VA.Drawing.Size SnapToNearestValue(VA.Drawing.Size size, VA.Drawing.Size snapsize)
        {
            return new VA.Drawing.Size(VA.Internal.MathUtil.Round(size.Width, snapsize.Width),
                            VA.Internal.MathUtil.Round(size.Height, snapsize.Height));
        }
    }
}