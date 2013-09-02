using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.Models.Charting
{
    public struct Sector
    {
        public double StartAngle { get; private set; }
        public double EndAngle { get; private set; }

        public Sector(double start, double end) :
            this()
        {
            this.StartAngle = start;
            this.EndAngle = end;
        }

        public double Angle
        {
            get { return this.EndAngle - this.StartAngle; }
        }
    }
}