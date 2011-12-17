using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Effects
{
    public class GradientStop
    {
        public Drawing.ColorRGB Color { get; set; }
        public VA.Drawing.Transparency Transparency { get; set; }
        public double Position { get; set; }

        public GradientStop(VA.Drawing.ColorRGB color, double transparency, double pos)
        {
            this.Color = color;
            this.Transparency = transparency;
            this.Position= pos;
        }
    }
}