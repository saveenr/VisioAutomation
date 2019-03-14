namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct NamedSrcValuePair
    {
        public readonly string Name;
        public readonly ShapeSheet.Src Src;
        public readonly string Value;

        public NamedSrcValuePair(string name, ShapeSheet.Src src, string value)
        {
            this.Name = name;
            this.Src = src;
            this.Value = value;
        }

        public static NamedSrcValuePair Create(string name, ShapeSheet.Src src, string value)
        {
            return new NamedSrcValuePair(name, src, value);
        }

        public static NamedSrcValuePair Create(string name, ShapeSheet.Src src, CellValueLiteral cvf)
        {
            return new NamedSrcValuePair(name, src, cvf.Value);
        }
    }

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

    public struct SidSrcValuePair
    {
        public readonly short ShapeID;
        public readonly ShapeSheet.Src Src;
        public readonly string Value;

        public SidSrcValuePair(short shapeid, ShapeSheet.Src src, string value)
        {
            this.ShapeID = shapeid;
            this.Src = src;
            this.Value = value;
        }

        public static SidSrcValuePair Create(short shapeid, ShapeSheet.Src src, string value)
        {
            return new SidSrcValuePair(shapeid, src, value);
        }

        public static SidSrcValuePair Create(short shapeid, ShapeSheet.Src src, CellValueLiteral cvf)
        {
            return new SidSrcValuePair(shapeid, src, cvf.Value);
        }
    }
}