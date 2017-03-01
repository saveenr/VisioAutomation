namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumn : ColumnBase
    {
        public ShapeSheet.SRC SRC { get; protected set; }

        internal CellColumn(int ordinal, ShapeSheet.SRC src, string name) :
            base(ordinal, name)
        {
            this.SRC = src;
        }

    }
}