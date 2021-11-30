namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SidSrcValue
    {
        public readonly short ShapeID;
        public readonly Core.Src Src;
        public readonly string Value;

        public SidSrcValue(short shapeid, Core.Src src, string value)
        {
            this.ShapeID = shapeid;
            this.Src = src;
            this.Value = value;
        }
    }
}
