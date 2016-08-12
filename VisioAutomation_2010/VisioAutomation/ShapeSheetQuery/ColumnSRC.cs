namespace VisioAutomation.ShapeSheetQuery
{
    public class ColumnSRC : ColumnBase
    {
        public ShapeSheet.SRC SRC { get; protected set; }

        internal ColumnSRC(int ordinal, ShapeSheet.SRC src, string name) :
            base(ordinal, name)
        {
            this.SRC = src;
        }

    }
}