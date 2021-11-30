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
    }
}