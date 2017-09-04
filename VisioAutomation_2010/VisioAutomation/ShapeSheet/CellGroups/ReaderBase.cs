using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderBase<TCellGroup>
    {
        protected ShapeSheetQuerySingle query;

        protected ReaderBase()
        {
            this.query = new ShapeSheetQuerySingle();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

    }

    public abstract class ReaderBaseMulti<TCellGroup>
    {
        protected ShapeSheetQuery query;

        protected ReaderBaseMulti()
        {
            this.query = new ShapeSheetQuery();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

    }

}