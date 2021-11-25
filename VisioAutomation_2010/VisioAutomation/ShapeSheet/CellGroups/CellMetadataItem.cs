namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct CellMetadataItem
    {
        public readonly string Name;
        public readonly VisioAutomation.Core.Src Src;
        public readonly string Value;

        public CellMetadataItem(string name, VisioAutomation.Core.Src src, string value)
        {
            this.Name = name;
            this.Src = src;
            this.Value = value;
        }
    }
}
