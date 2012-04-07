using Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Layout.Models.Pie
{
    public class PieSlice
    {
        public double Value { get; set; }
        public Shape VisioShape { get; set; }
        public object Data { get; set; }
        public string Text { get; set; }
        public double EndAngle { get; set; }
        public double StartAngle { get; set; }
    }
}