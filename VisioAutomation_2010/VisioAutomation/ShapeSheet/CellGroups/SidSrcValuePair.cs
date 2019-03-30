namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SidSrcValuePair
    {
        public readonly short ShapeID;
        public readonly ShapeSheet.Src Src;
        public readonly string Value;

        public SidSrcValuePair(short shapeid, ShapeSheet.Src src, string value)
        {
            this.ShapeID = shapeid;
            this.Src = src;
            this.Value = value;
        }

        public static SidSrcValuePair Create(short shapeid, ShapeSheet.Src src, string value)
        {
            return new SidSrcValuePair(shapeid, src, value);
        }

        public static SidSrcValuePair Create(short shapeid, ShapeSheet.Src src, CellValueLiteral cvf)
        {
            return new SidSrcValuePair(shapeid, src, cvf.Value);
        }
    }
}
