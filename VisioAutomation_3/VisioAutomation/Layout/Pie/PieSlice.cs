using Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Layout.Pie
{
    public class PieSlice
    {
        public double Value { get; set; }
        public Shape VisioShape { get; set; }
        public object Data { get; set; }
        public string Text { get; set; }
        public VA.Angle EndAngle { get; set; }
        public VA.Angle StartAngle { get; set; }
    }
}