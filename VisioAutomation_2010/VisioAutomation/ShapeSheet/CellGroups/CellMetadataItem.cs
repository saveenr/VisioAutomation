namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct CellMetadataItem
    {
        public readonly string Name;
        public readonly ShapeSheet.Src Src;
        public readonly string Value;

        public CellMetadataItem(string name, ShapeSheet.Src src, string value)
        {
            this.Name = name;
            this.Src = src;
            this.Value = value;
        }

        public static CellMetadataItem Create(string name, ShapeSheet.Src src, string value)
        {
            return new CellMetadataItem(name, src, value);
        }

        public static CellMetadataItem Create(string name, ShapeSheet.Src src, CellValueLiteral cvf)
        {
            return new CellMetadataItem(name, src, cvf.Value);
        }
    }
}
