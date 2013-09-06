using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Models.Charting
{
    public class DataPoint
    {
        public double Value;
        public string Label;
        public string LabelFormat;
        public IVisio.Shape VisioShape;

        public DataPoint()
        {
            
        }

        public DataPoint(double value)
        {
            this.Value = value;
        }
    }
}