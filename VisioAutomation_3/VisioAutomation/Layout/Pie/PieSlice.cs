using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.Pie
{
    public class PieSlice
    {
        public double Value { get; set; }
        public Shape VisioShape { get; set; }
        public object Data { get; set; }
        public string Text { get; set; }
    }
}