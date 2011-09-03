using IVisio = Microsoft.Office.Interop.Visio;

namespace VisUtil
{
    public static class DrawUtil
    {
        public static IVisio.Shape DrawCircleFromCenter(IVisio.Page page, double x, double y, double r)
        {
            var shape = page.DrawOval(x - r, y - r, x + r, y + r);
            return shape;
        }
    }
}
