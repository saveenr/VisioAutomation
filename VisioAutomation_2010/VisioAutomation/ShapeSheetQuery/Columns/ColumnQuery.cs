namespace VisioAutomation.ShapeSheetQuery.Columns
{
    public class ColumnQuery : ColumnBase
    {
        public ShapeSheet.SRC SRC { get; protected set; }

        internal ColumnQuery(int ordinal, ShapeSheet.SRC src, string name) :
            base(ordinal, name)
        {
            this.SRC = src;
        }

    }
}