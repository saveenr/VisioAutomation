using VA = VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public struct TabStop
    {
        public double Position { get; private set; }
        public TabStopAlignment Alignment { get; private set; }

        public TabStop(double pos, VA.Text.TabStopAlignment align) : this()
        {
            this.Position = pos;
            this.Alignment = align;
        }

        public override string ToString()
        {
            string s = string.Format("(Position={0},Alignment={1})",
                                     this.Position,
                                     this.Alignment);
            return s;
        }
    }
}