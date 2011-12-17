using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static partial class ShapeMethods
    {
        public static IVisio.Cell GetCell(this IVisio.Shape shape, VA.ShapeSheet.SRC src)
        {
            return shape.CellsSRC[src.Section, src.Row, src.Cell];
        }

        public static VA.Drawing.Rectangle GetBoundingBox(this IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            // MSDN: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vissdk11/html/vimthBoundingBox_HV81900422.asp
            double bbx0, bby0, bbx1, bby1;
            shape.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new VA.Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static VA.Drawing.Point XYFromPage(this IVisio.Shape shape, VA.Drawing.Point xy)
        {
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new VA.Drawing.Point(xprime, yprime);
        }

        public static VA.Drawing.Point XYToPage(this IVisio.Shape shape, VA.Drawing.Point xy)
        {
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new VA.Drawing.Point(xprime, yprime);
        }
    }
}