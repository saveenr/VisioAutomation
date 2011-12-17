using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class ColorMethods
    {
        public static VA.Drawing.ColorRGB ToColorRGB(this IVisio.Color color)
        {
            return new VA.Drawing.ColorRGB(color.Red, color.Green, color.Blue);
        }
    }
}