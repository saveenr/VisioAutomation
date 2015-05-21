using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public struct TabStop
    {
        public double Position { get; private set; }
        public TabStopAlignment Alignment { get; private set; }

        public TabStop(double pos, TabStopAlignment align) : this()
        {
            this.Position = pos;
            this.Alignment = align;
        }

        public override string ToString()
        {
            string s = $"(Position={this.Position},Alignment={this.Alignment})";
            return s;
        }
    }
}