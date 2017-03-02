namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumn : ColumnBase
    {
        public ShapeSheet.Src Src { get; protected set; }

        internal CellColumn(int ordinal, ShapeSheet.Src src, string name) :
            base(ordinal, name)
        {
            this.Src = src;
        }

    }
}