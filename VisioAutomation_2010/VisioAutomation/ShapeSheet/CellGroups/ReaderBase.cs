using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderBase<TCellGroup>
    {
        protected ShapeSheetQuery query;

        protected ReaderBase()
        {
            this.query = new ShapeSheetQuery();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(ShapeSheet.CellData[] row);

    }
}