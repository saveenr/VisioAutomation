namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcValuePair
    {
        public readonly ShapeSheet.Src Src;
        public readonly string Value;

        public SrcValuePair(ShapeSheet.Src src, string value)
        {
            this.Src = src;
            this.Value = value;
        }

        public static SrcValuePair Create(ShapeSheet.Src src, string value)
        {
            return new SrcValuePair(src,value);
        }

        public static SrcValuePair Create(ShapeSheet.Src src, CellValueLiteral cvf)
        {
            return new SrcValuePair(src, cvf.Value);
        }
    }
}