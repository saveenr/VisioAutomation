using VisioAutomation.ShapeSheet.Queries;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal struct QueryCellInfo
    {
        public SIDSRC SIDSRC;
        public ColumnBase Column;

        public QueryCellInfo(SIDSRC sidsrc, ColumnBase col)
        {
            this.SIDSRC = sidsrc;
            this.Column = col;
        }
    }
}