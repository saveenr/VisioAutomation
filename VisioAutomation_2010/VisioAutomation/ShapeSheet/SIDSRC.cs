namespace VisioAutomation.ShapeSheet
{
    public struct SidSrc
    {
        public short ShapeID { get; }
        public Src Src { get; }

        public SidSrc(
            short shapeid,
            Src src)
        {
            this.ShapeID = shapeid;
            this.Src = src;
        }

        public SidSrc(
            short shapeid,
            short section,
            short row,
            short cell)
        {
            this.ShapeID = shapeid;
            this.Src = new Src(section,row,cell);
        }

        public override string ToString()
        {
            return string.Format("{0}({1},{2},{3},{4})", nameof(SidSrc),this.ShapeID, this.Src.Section, this.Src.Row, this.Src.Cell);
        }
    }
}