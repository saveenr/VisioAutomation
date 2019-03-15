namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumn : ColumnBase
    {
        public readonly ShapeSheet.Src Src;

        internal CellColumn(int ordinal, ShapeSheet.Src src, string name) :
            base(ordinal, name)
        {
            this.Src = src;
        }

    }
}