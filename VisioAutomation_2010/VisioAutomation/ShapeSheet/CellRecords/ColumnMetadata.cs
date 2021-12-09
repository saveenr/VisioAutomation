namespace VisioAutomation.ShapeSheet.CellRecords
{
    public readonly struct ColumnMetadata
    {
        public readonly string Name;
        public readonly Core.Src Src;
        public readonly string Value;

        public ColumnMetadata(string name, Core.Src src, string value)
        {
            this.Name = name;
            this.Src = src;
            this.Value = value;
        }
    }
}