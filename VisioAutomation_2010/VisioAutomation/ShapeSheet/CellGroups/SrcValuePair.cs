namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcValuePair
    {
        public readonly Core.Src Src;
        public readonly string Value;

        public SrcValuePair(Core.Src src, string value)
        {
            this.Src = src;
            this.Value = value;
        }

        public static SrcValuePair Create(Core.Src src, string value)
        {
            return new SrcValuePair(src,value);
        }

        public static SrcValuePair Create(Core.Src src, Core.CellValue cvf)
        {
            return new SrcValuePair(src, cvf.Value);
        }
    }
}