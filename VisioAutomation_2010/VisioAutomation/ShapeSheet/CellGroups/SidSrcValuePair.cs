namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SidSrcValuePair
    {
        public readonly short ShapeID;
        public readonly VisioAutomation.Core.Src Src;
        public readonly string Value;

        public SidSrcValuePair(short shapeid, VisioAutomation.Core.Src src, string value)
        {
            this.ShapeID = shapeid;
            this.Src = src;
            this.Value = value;
        }

        public static SidSrcValuePair Create(short shapeid, VisioAutomation.Core.Src src, string value)
        {
            return new SidSrcValuePair(shapeid, src, value);
        }

        public static SidSrcValuePair Create(short shapeid, VisioAutomation.Core.Src src, VisioAutomation.Core.CellValue cvf)
        {
            return new SidSrcValuePair(shapeid, src, cvf.Value);
        }
    }
}
