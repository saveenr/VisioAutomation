namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct CellMetadataItem
    {
        public readonly string Name;
        public readonly Core.Src Src;
        public readonly string Value;

        public CellMetadataItem(string name, Core.Src src, string value)
        {
            this.Name = name;
            this.Src = src;
            this.Value = value;
        }
    }
}
