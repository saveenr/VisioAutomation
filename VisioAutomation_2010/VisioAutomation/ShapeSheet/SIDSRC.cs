namespace VisioAutomation.ShapeSheet
{
    public struct SidSrc
    {
        public short ShapeID { get; }
        public Src Src { get; }

        public SidSrc(
            short shape_id,
            Src src)
        {
            this.ShapeID = shape_id;
            this.Src = src;
        }

        public SidSrc(
            short shape_id,
            short section,
            short row,
            short cell)
        {
            this.ShapeID = shape_id;
            this.Src = new Src(section,row,cell);
        }

        public override string ToString()
        {
            return string.Format("{0}({1},{2},{3},{4})", nameof(SidSrc),this.ShapeID, this.Src.Section, this.Src.Row, this.Src.Cell);
        }
    }
}