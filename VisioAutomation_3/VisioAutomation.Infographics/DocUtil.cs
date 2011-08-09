using VA=VisioAutomation;

namespace VisioAutomation.Infographics
{
    public static class DocUtil
    {
        public static VA.Drawing.Rectangle BuildFromUpperLeft(VA.Drawing.Point upperleft, VA.Drawing.Size s)
        {
            var rect = new VA.Drawing.Rectangle(upperleft.X, upperleft.Y - s.Height, upperleft.X + s.Width, upperleft.Y);
            return rect;
        }
    }
}