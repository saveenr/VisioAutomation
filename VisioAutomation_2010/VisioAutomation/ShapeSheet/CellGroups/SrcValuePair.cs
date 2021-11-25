namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcValuePair
    {
        public readonly VisioAutomation.Core.Src Src;
        public readonly string Value;

        public SrcValuePair(VisioAutomation.Core.Src src, string value)
        {
            this.Src = src;
            this.Value = value;
        }

        public static SrcValuePair Create(VisioAutomation.Core.Src src, string value)
        {
            return new SrcValuePair(src,value);
        }

        public static SrcValuePair Create(VisioAutomation.Core.Src src, VisioAutomation.Core.CellValue cvf)
        {
            return new SrcValuePair(src, cvf.Value);
        }
    }
}