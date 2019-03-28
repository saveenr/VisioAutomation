namespace VisioAutomation.Text
{
    public struct TabStop
    {
        public double Position { get; }
        public TabStopAlignment Alignment { get; }

        public TabStop(double pos, TabStopAlignment align) : this()
        {
            this.Position = pos;
            this.Alignment = align;
        }

        public override string ToString()
        {
            string s = string.Format("(Position={0},Alignment={1})", this.Position, this.Alignment);
            return s;
        }
    }
}