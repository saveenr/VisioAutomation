namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct CellMetadata
    {
        public readonly string Name;
        public readonly Core.Src Src;
        public readonly string Value;

        public CellMetadata(string name, Core.Src src, string value)
        {
            this.Name = name;
            this.Src = src;
            this.Value = value;
        }
    }
}
