namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumn : ColumnBase
    {
        public ShapeSheet.Src SRC { get; protected set; }

        internal CellColumn(int ordinal, ShapeSheet.Src src, string name) :
            base(ordinal, name)
        {
            this.SRC = src;
        }

    }
}