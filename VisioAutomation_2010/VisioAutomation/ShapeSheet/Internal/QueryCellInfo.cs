using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal struct QueryCellInfo
    {
        public SidSrc SidSrc;
        public Column Column;

        public QueryCellInfo(SidSrc sidsrc, Column col)
        {
            this.SidSrc = sidsrc;
            this.Column = col;
        }
    }
}