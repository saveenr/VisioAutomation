namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcValuePair
    {
        public readonly ShapeSheet.Src Src;
        public readonly string Formula;

        public SrcValuePair(ShapeSheet.Src src, string formula)
        {
            this.Src = src;
            this.Formula = formula;
        }

        public static SrcValuePair Create(ShapeSheet.Src src, string formula)
        {
            return new SrcValuePair(src,formula);
        }
    }
}