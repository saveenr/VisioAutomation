namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcValue
    {
        public readonly Core.Src Src;
        public readonly string Value;

        public SrcValue(Core.Src src, string value)
        {
            this.Src = src;
            this.Value = value;
        }
    }
}