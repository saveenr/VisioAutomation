namespace VisioAutomation.ShapeSheet.Queries
{
    public class ColumnCell : ColumnBase
    {
        public ShapeSheet.SRC SRC { get; protected set; }

        internal ColumnCell(int ordinal, ShapeSheet.SRC src, string name) :
            base(ordinal, name)
        {
            this.SRC = src;
        }

    }
}