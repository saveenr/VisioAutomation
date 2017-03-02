using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal struct QueryCellInfo
    {
        public SidSrc SidSrc;
        public ColumnBase Column;

        public QueryCellInfo(SidSrc sidsrc, ColumnBase col)
        {
            this.SidSrc = sidsrc;
            this.Column = col;
        }
    }
}